import type { IFunctionInfo } from '@univerjs/engine-formula'
import { useEffect, useMemo, useRef, useState } from 'react'
import { platformShortcuts } from '@genoffice/i18n'
import {
  Dropdown,
  SHAPE_GALLERY_GROUPS,
  ShapePreview,
  useDismissablePopover,
  useRibbonCollapse,
  aiPanelInitiallyOpen,
  rememberAiPanelOpen,
} from '@genoffice/ui'

import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderOuterIcon,
  BorderRightIcon,
  BorderThickOuterIcon,
  BorderTopIcon,
  CaretIcon,
  GensparkMark,
  RIBBON_GLYPH_ICONS,
  RedoIcon,
  SaveAsIcon,
  SaveIcon,
  UndoIcon,
} from './ribbon-icons'
import { sheetContextMenuAction } from './sheet-context-menu'

import { ColorDropdown } from './ColorDropdown'
import { FormatCellsDialog } from './FormatCellsDialog'
import { AllowEditRangesDialog } from './AllowEditRangesDialog'
import { ProtectSheetDialog } from './ProtectSheetDialog'
import type { SheetProtectionAllow } from '@genoffice/xlsx-gateway/gateway/xlsx-protection'
import { GoToDialog } from './GoToDialog'
import { SheetTabContextMenu } from './SheetTabContextMenu'
import type { SheetTabActions } from './sheet-tab-actions'
import { PasteSpecialDialog } from './PasteSpecialDialog'
import {
  PASTE_SPECIAL_OPEN_SHELL_COMMAND,
  type PasteSpecialOptions,
  type PasteSpecialSource,
} from './paste-special'
import { COLOR_SCHEMES, FONT_SCHEMES, THEME_PRESETS } from './themes'
import { useI18n, type StringKey } from './i18n/locale'
import { NameManagerDialog, type DefinedNameAction, type DefinedNameRow } from './NameManagerDialog'
import {
  NUMBER_FORMAT_LABEL,
  categoryOptionForPattern,
  numberFormatCategories,
} from './number-format'
import { setContextMenuCommandSink, registerPasteSpecialOpener } from './cell-context-menu'
import { numberFormatLabel, type SelectionFormat } from './selection-format'
import { fontFamilyGroups, useSystemFontFamilies } from './system-fonts'
import {
  isFormulaBarKeyTarget,
  isGridKeyTarget,
  shouldInterceptClearSelection,
} from './clear-selection-keyboard'
import {
  accountingStyle,
  COMMA_STYLE,
  PERCENT_STYLE,
  RIBBON_CURRENCY_SYMBOL,
} from './excel-format-shortcuts'
import { isModalOpen, resolveGlobalShortcut } from './global-shortcuts'
import { stepFontSize } from './font-size-ladder'

import type { ChartSeriesVisualState } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { ChangePlan } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { AttachmentMeta } from '../shared/desktop-api'
import { AiChatPanel, type AiChatMessage } from './ai/AiChatPanel'
import { AiSelectionAsk } from './ai/AiSelectionAsk'
import type { SelectionAskAnchor } from './ai/selection-ask'
import {
  PivotDialog,
  type PivotCreateRequest,
  type PivotField,
  type OoXmlPivotConfig,
} from './PivotDialog'
import { PivotFieldPane } from './PivotFieldPane'
import type { PivotEditSeed } from './pivot-actions'
import {
  modelToConfig,
  validatePivotModel,
  type PivotLayoutModel,
  type PivotModelError,
} from './pivot-field-model'
import type { GoalSeekResult } from './goal-seek'
import { GoalSeekDialog } from './GoalSeekDialog'
import type { FillSeriesOptions, FillSeriesType } from './fill-series'
import type { FillSeriesContext } from './fill-series-apply'
import { FillSeriesDialog } from './FillSeriesDialog'
import {
  registerFxButtonOpener,
  type FunctionArgumentsHost,
  type FunctionTarget,
} from './function-arguments-runtime'
import { InsertFunctionFlow } from './InsertFunctionFlow'
import { SubtotalDialog, type SubtotalConfig } from './SubtotalDialog'
import { ConsolidateDialog } from './ConsolidateDialog'
import type { ConsolidateConfig } from './consolidate'
import { HeaderFooterDialog, type HeaderFooterResult } from './HeaderFooterDialog'
import type { PrintPreviewHost } from './page-layout-actions'
import { PrintDialog } from './PrintDialog'
import type { HeaderFooterParts } from './edit-journal'
import { useModalDialog } from './modal-dialog'
import type { SelectedTableRibbon } from './table-design-actions'
import type { TableOptionKey } from './table-design'
import { builtinTablePalette, tableStyleNames, type TableStyleFamily } from './table-styles'
import { STATUS_BAR_FUNCS, type StatusBarFunc } from './status-bar-stats'
import { SHEET_ZOOM_MAX, SHEET_ZOOM_MIN, clampZoomPercent } from './zoom-range'
import { ZoomDialog } from './ZoomDialog'

// No File tab: file commands live in the macOS
// application menu (File → Open/Save/Save As) and the toolbar icons.
const ribbonTabs = ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'] as const

/// 'Chart Design' is contextual: it exists only while a chart is selected,
/// and appears without stealing the active tab.
type RibbonTab =
  (typeof ribbonTabs)[number] | 'Chart Design' | 'Table Design' | 'Shape Format' | 'Picture Format'
const CONTEXTUAL_TABS: readonly RibbonTab[] = [
  'Chart Design',
  'Table Design',
  'Shape Format',
  'Picture Format',
]

// Internal values keep English ids (=== matching, state values unchanged);
// translated only for display
const TAB_LABEL: Record<RibbonTab, StringKey> = {
  Home: 'appTabHome',
  Insert: 'appTabInsert',
  'Page Layout': 'appTabPageLayout',
  Formulas: 'appTabFormulas',
  Data: 'appTabData',
  Review: 'appTabReview',
  View: 'appTabView',
  'Chart Design': 'appTabChartDesign',
  'Shape Format': 'appTabShapeFormat',
  'Table Design': 'appTabTableDesign',
  'Picture Format': 'appTabPictureFormat',
}

/// Paint of the selected editable shape; `'none'` is an explicit no fill/line.
export interface SelectedShapeRibbon {
  readonly fillColor: string | undefined
  readonly lineColor: string | undefined
}

export interface SelectedChartRibbon {
  readonly title: string
  /// Non-null when the chart's type can be rewritten (single-plot
  /// column/bar/line/area/pie/doughnut); scatter/combo/3D keep their type.
  readonly convertible: string | null
  readonly currentType: string | null
  readonly canEdit: boolean
  readonly isPie: boolean
  readonly hasAxes: boolean
  /// Whether the save pipeline can write plot-level data labels for this
  /// chart family (scatter/radar cannot).
  readonly canLabel: boolean
  readonly seriesCount: number
  readonly categoryCount: number
  /// Current series (pending edits applied) for Switch Row/Column.
  readonly series: readonly ChartSeriesVisualState[]
  readonly legend?: string | undefined
  readonly axisTitles?:
    | {
        category?: string | null | undefined
        value?: string | null | undefined
      }
    | undefined
  readonly dataLabels?: string | undefined
  readonly grouping?: string | undefined
}

type ChartTextTarget = 'title' | 'axis-category' | 'axis-value'

const CHART_TEXT_LABELS: Record<ChartTextTarget, { heading: StringKey; command: string }> = {
  title: { heading: 'appChartTitleEl', command: 'chart-title' },
  'axis-category': { heading: 'appCategoryAxisTitle', command: 'chart-axis-cat' },
  'axis-value': { heading: 'appValueAxisTitle', command: 'chart-axis-val' },
}

/// mode=tab: embedded in the shell's tab strip, which owns the traffic
/// lights / caption buttons — the ribbon must not reserve space for them.
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'
const IS_MAC = navigator.platform.toLowerCase().includes('mac')

/// Univer's formula bar keeps its expanded state in React component state
/// with no command behind it; the fold arrow is the only way in.
function toggleFormulaBarExpand(): void {
  document
    .querySelector<HTMLElement>('[data-u-comp="formula-bar"] .univer-w-5.univer-cursor-pointer')
    ?.click()
}

/// Review > Translate targets, shown in their native names (never localized).
const TRANSLATE_LANGUAGES = [
  'English',
  '简体中文',
  '繁體中文',
  '日本語',
  '한국어',
  'Español',
  'Français',
  'Deutsch',
  'Português',
  'Русский',
  'العربية',
] as const

/// Glyph icons live in ribbon-icons.tsx, drawn to the shared icon standard
/// (24×24 canvas, 1.5px strokes, round caps/joins); unmapped glyphs render
/// as text (letterforms such as $, ?, θ, ƒx are typography, not icons).
function ToolSymbol({ symbol }: { readonly symbol: string }): React.JSX.Element {
  return (
    <span className="tool-symbol" aria-hidden="true">
      {RIBBON_GLYPH_ICONS[symbol] ?? symbol}
    </span>
  )
}

interface ExcelShellProps {
  readonly openingWorkbook: boolean
  readonly prompt: string
  readonly preview: ChangePlan | null
  readonly selectionFormat: SelectionFormat | null
  readonly formatPainterActive: boolean
  /// True when the workbook has any cell content (the one-click AI action
  /// buttons are greyed out on an empty sheet).
  readonly sheetHasContent: boolean
  /// true while the real LLM agent is running (composer disabled meanwhile).
  readonly aiBusy: boolean
  readonly chat: readonly AiChatMessage[]
  readonly historicChat?: readonly AiChatMessage[]
  /// Chat attachments (chips + 📎 button + drag-and-drop), same structure as the
  /// docs/slides AI panels.
  readonly attachments: readonly AttachmentMeta[]
  readonly attachNotice: string | null
  readonly onPickAttachments: () => void
  readonly onAddAttachmentPaths: (paths: readonly string[]) => void
  readonly onAddPastedImage: (data: ArrayBuffer, ext: string) => void
  readonly onRemoveAttachment: (path: string) => void
  readonly onPromptChange: (prompt: string) => void
  /** Send the composer text, or the given instruction when provided (Retry also
   *  resends that message's original attachments and passes the failed bubble's
   *  chat index so the send replaces it in place) */
  readonly onSend: (
    instruction?: string,
    attachments?: readonly AttachmentMeta[],
    retryIndex?: number,
  ) => void
  readonly onStop: () => void
  readonly onNewChat: () => void
  readonly onUndo: (steps?: number) => void
  /// A1 notation of the multi-cell selection the AI composer offers as this
  /// run's scope, or null when the resting single-cell selection carries none.
  readonly aiScopeRange: string | null
  /// Header names when that scope covers whole columns: they label the chip in
  /// place of the range, because a column is a name to the user, not a letter.
  readonly aiScopeColumns: readonly string[] | null
  /// The range above belongs to a run in flight and can no longer be dropped.
  readonly aiScopeLocked: boolean
  /// Drag endpoint and viewport bounds used to place the localized trigger.
  readonly aiSelectionAskAnchor: SelectionAskAnchor | null
  readonly onAiSelectionAskDismiss: () => void
  readonly onAiScopeDismiss: () => void
  /// Citation link in an AI answer: jumps the grid to the cited cell/range.
  readonly onAiCitation: (href: string) => void
  readonly onCommand: (command: string) => void
  /// True while Univer's in-cell editor is open (Backspace must delete
  /// characters, not clear the selection).
  readonly onIsCellEditing: () => boolean
  /// Left side of the status bar (ready / streaming / AI progress messages).
  readonly statusMessage: string
  readonly emptyCsvNotice: boolean
  readonly onOpenWorkbook: () => void
  readonly onDismissEmptyCsvNotice: () => void
  /// Zoom of the active sheet in percent, echoed by the status-bar slider.
  readonly zoomPercent: number
  /// Aggregates ticked in the status bar's right-click menu (Excel parity).
  readonly statusBarFuncs: readonly StatusBarFunc[]
  readonly onToggleStatusBarFunc: (func: StatusBarFunc) => void
  /// True when the edit journal has unsaved changes (enables the QAT Save).
  readonly canSave: boolean
  readonly onSave: () => void
  /// Save As remains available for a clean workbook, but requires a real
  /// file-backed session (the in-memory demo workbook has nowhere to copy).
  readonly canSaveAs: boolean
  readonly onSaveAs: () => void
  /// QAT redo (workbook history, same path as the app menu's ⇧⌘Z); undo
  /// shares the AI panel's onUndo above.
  readonly onRedo: () => void
  /// Undo/redo stack occupancy: the QAT buttons grey out when there is nothing to apply.
  readonly canUndo: boolean
  readonly canRedo: boolean
  /// AutoSave toggle in the tab row (docs/slides parity).
  readonly autoSave: boolean
  readonly onAutoSaveChange: (on: boolean) => void
  /// Non-null while a floating chart is selected in the grid.
  readonly selectedChart: SelectedChartRibbon | null
  /// Non-null while an editable shape is selected (Shape Format tab).
  readonly selectedShape: SelectedShapeRibbon | null
  /// Kind of the selected floating picture/shape (Shape/Picture Format tab).
  readonly selectedVisualKind: 'image' | 'shape' | null
  /// Non-null while the selection sits inside a table (Table Design tab).
  readonly selectedTable: SelectedTableRibbon | null
  /// Active sheet's outlinePr echo for the Data > Group settings toggles.
  readonly outlineSummary: { below: boolean; right: boolean }
  /// Bumps after each successful Create Table; the Table Design tab activates.
  readonly tableCreatedSeq: number
  /// Default range for the Ctrl+T dialog (selection widened to its data block).
  readonly onGetCreateTableRange: () => string
  /// Column choices of the active selection, read when the Sort dialog opens.
  readonly onGetSortColumns: () => { label: string; colIndex: number }[]
  /// Effective protection of the active sheet (null = unknown / demo).
  readonly onGetSheetProtection: () => { protected: boolean; hasPassword: boolean } | null
  /// Protect Sheet dialog result; resolve to an error message or null.
  readonly onProtectSheet: (request: {
    password: string | null
    allow: SheetProtectionAllow
  }) => Promise<string | null>
  readonly onUnprotectSheet: (password: string) => Promise<string | null>
  readonly sheetTabActions: SheetTabActions | null
  /// Effective workbook structure lock (null = no file open).
  readonly onGetWorkbookProtection: () => boolean | null
  readonly formulaBarVisible: boolean
  /// Cross-highlight ("reading mode") of the active row/column, echoed by the View checkbox.
  readonly crossHighlightVisible: boolean
  /// Allow-edit ranges of the active sheet, read when the dialog opens.
  readonly onGetProtectedRanges: () => {
    ranges: readonly { name: string; sqref: string; hasPassword: boolean }[]
    error: string | null
  }
  readonly onApplyProtectedRanges: (
    ranges: readonly { name: string; sqref: string }[],
  ) => string | null
  /// Name Manager data + actions (actions return an error message or null).
  readonly onGetDefinedNames: () => {
    names: DefinedNameRow[]
    sheets: { id: string; name: string }[]
    activeSheetId: string | null
  }
  readonly onDefinedNameAction: (action: DefinedNameAction) => string | null
  /// Subtotals use the selection; pivots pass their resolved source range.
  readonly onGetPivotFields: (sourceRange?: string) => PivotField[]
  readonly onGetSourceRange: () => string
  readonly onCreatePivot: (request: PivotCreateRequest) => string | null
  /// PivotTable Fields pane: seed of the pivot under the cursor (null = App
  /// already showed why), the apply callback, and the member list of one
  /// source field for the report-filter item picker.
  readonly onGetPivotEditSeed: () => PivotEditSeed | null
  readonly onEditPivot: (config: OoXmlPivotConfig) => string | null
  readonly onGetPivotFieldMembers: (fieldIndex: number) => string[]
  /// Identity of the pivot under the selection; the pane follows it like Excel.
  readonly pivotSelectionKey: string | null
  readonly onRefreshPivot: () => string | null
  readonly onIsSelectionInPivot: () => boolean
  readonly onGetActiveCell: () => string
  /// Value of the selection's top-left cell, read when Format Cells opens
  /// (number-format preview).
  readonly onGetAnchorValue: () => number | string | null
  /// A1 label of the active cell, echoed live by the Name Box.
  readonly activeCellA1: string
  /// Name Box / Go To jump; returns an error message, or null on success.
  readonly onGoToReference: (ref: string) => string | null
  readonly onListDefinedNames: () => readonly { name: string; ref: string }[]
  readonly onReadPasteSpecialSource: () => Promise<PasteSpecialSource>
  /// Resolves to an error message, or null on success.
  readonly onPasteSpecial: (
    source: PasteSpecialSource,
    options: PasteSpecialOptions,
  ) => Promise<string | null>
  /// Function descriptions from the running formula engine (Insert Function).
  readonly onListFunctions: () => readonly IFunctionInfo[]
  readonly functionArguments: FunctionArgumentsHost
  readonly onCreateSubtotal: (config: SubtotalConfig) => string | null
  readonly onCreateConsolidate: (config: ConsolidateConfig) => string | null
  /// Prefill for the Consolidate reference input (current multi-cell selection).
  readonly onGetConsolidateDefault: () => string
  /// Header & Footer dialog OK; returns an error message, or null on success.
  readonly onApplyHeaderFooter: (result: HeaderFooterResult) => string | null
  /// Print with Preview, open while non-null.
  readonly printHost: PrintPreviewHost | null
  readonly onClosePrint: () => void
  /// Session page-layout settings of the active sheet, echoed by the Page
  /// Layout tab's controls (untouched fields show the app default).
  readonly pageLayout: PageLayoutEcho
  /// Manual-recalc mode echo for the Calculation Options menu.
  readonly calcManual: boolean
  /// Goal Seek solve; rejects with a user-facing Error message.
  readonly onGoalSeek: (setCell: string, toValue: number, byCell: string) => Promise<GoalSeekResult>
  /// Fill > Series… OK; returns an error message, or null on success.
  readonly onFillSeries: (options: FillSeriesOptions) => string | null
  /// Series dialog defaults (selection shape, workbook date system).
  readonly onGetFillSeriesContext: () => FillSeriesContext
}

export interface PageLayoutEcho {
  readonly orientation?: 'portrait' | 'landscape' | undefined
  readonly paperSize?: number | undefined
  readonly scale?: number | undefined
  readonly fitToWidth?: number | undefined
  readonly fitToHeight?: number | undefined
  readonly margins?: 'normal' | 'wide' | 'narrow' | undefined
  readonly printGridlines?: boolean | undefined
  readonly printHeadings?: boolean | undefined
  readonly showGridlines: boolean
  readonly showHeadings: boolean
  readonly printArea?: string | null | undefined
  readonly printTitles?: string | null | undefined
  readonly header?: HeaderFooterParts | null | undefined
  readonly footer?: HeaderFooterParts | null | undefined
  /// Page Break Preview overlay on for the active sheet (View tab echo).
  readonly pageBreakPreview?: boolean | undefined
}

interface PivotPaneState {
  readonly seed: PivotEditSeed
  readonly model: PivotLayoutModel
  readonly deferred: boolean
  readonly dirty: boolean
  readonly error: string | null
}

const PIVOT_MODEL_ERRORS: Record<PivotModelError, StringKey> = {
  needRow: 'appNeedRowField',
  needValue: 'dlgPivotErrNeedValue',
  colNeedsSingleValue: 'dlgPivotErrColNeedsSingleValue',
  calcIncomplete: 'dlgPivotErrCalcFieldIncomplete',
  rangeStep: 'dlgPivotErrRangeStep',
  labelFilterText: 'dlgPivotErrLabelFilterText',
  oneValueFilter: 'dlgPivotErrOneValueFilter',
  filterConflict: 'dlgPivotErrFilterConflict',
  dupValueName: 'dlgPivotErrDupValueName',
}

export function ExcelShell({
  prompt,
  preview,
  selectionFormat,
  formatPainterActive,
  sheetHasContent,
  aiBusy,
  chat,
  historicChat,
  attachments,
  attachNotice,
  onPickAttachments,
  onAddAttachmentPaths,
  onAddPastedImage,
  onRemoveAttachment,
  onGetSortColumns,
  onGetSheetProtection,
  onProtectSheet,
  onUnprotectSheet,
  sheetTabActions,
  onGetWorkbookProtection,
  formulaBarVisible,
  crossHighlightVisible,
  onGetProtectedRanges,
  onApplyProtectedRanges,
  onGetDefinedNames,
  onDefinedNameAction,
  onGetPivotFields,
  onGetSourceRange,
  onCreatePivot,
  onGetPivotEditSeed,
  onEditPivot,
  onGetPivotFieldMembers,
  pivotSelectionKey,
  onRefreshPivot,
  onIsSelectionInPivot,
  onGetActiveCell,
  onGetAnchorValue,
  activeCellA1,
  onGoToReference,
  onListDefinedNames,
  onReadPasteSpecialSource,
  onPasteSpecial,
  functionArguments,
  onListFunctions,
  onCreateSubtotal,
  onCreateConsolidate,
  onGetConsolidateDefault,
  onApplyHeaderFooter,
  printHost,
  onClosePrint,
  onPromptChange,
  onSend,
  onStop,
  onNewChat,
  onUndo,
  aiScopeRange,
  aiScopeColumns,
  aiScopeLocked,
  aiSelectionAskAnchor,
  onAiSelectionAskDismiss,
  onAiScopeDismiss,
  onAiCitation,
  onCommand,
  onIsCellEditing,
  openingWorkbook,
  statusMessage,
  emptyCsvNotice,
  onOpenWorkbook,
  onDismissEmptyCsvNotice,
  zoomPercent,
  statusBarFuncs,
  onToggleStatusBarFunc,
  canSave,
  onSave,
  canSaveAs,
  onSaveAs,
  onRedo,
  canUndo,
  canRedo,
  autoSave,
  onAutoSaveChange,
  selectedChart,
  selectedShape,
  selectedTable,
  outlineSummary,
  selectedVisualKind,
  tableCreatedSeq,
  onGetCreateTableRange,
  pageLayout,
  calcManual,
  onGoalSeek,
  onFillSeries,
  onGetFillSeriesContext,
}: ExcelShellProps): React.JSX.Element {
  const { t } = useI18n()
  const [activeTab, setActiveTab] = useState<RibbonTab>('Home')
  const collapse = useRibbonCollapse('ai-sheets-ribbon-collapsed', {
    collapse: t('appRibbonCollapse'),
    expand: t('appRibbonExpand'),
  })
  // Persisted so a closed AI panel stays closed on next launch (docs/slides parity)
  const [isCopilotOpen, setIsCopilotOpen] = useState(() =>
    aiPanelInitiallyOpen('ai-sheets-show-ai'),
  )
  useEffect(() => {
    rememberAiPanelOpen('ai-sheets-show-ai', isCopilotOpen)
  }, [isCopilotOpen])
  const [showFormatCells, setShowFormatCells] = useState(false)
  const [axisSizeTarget, setAxisSizeTarget] = useState<'row' | 'col' | null>(null)
  const [showLinkDialog, setShowLinkDialog] = useState(false)
  const [showSortDialog, setShowSortDialog] = useState(false)
  const [showDedupeDialog, setShowDedupeDialog] = useState(false)
  const [showCreateTable, setShowCreateTable] = useState(false)
  const [showResizeTable, setShowResizeTable] = useState(false)
  const [showNameManager, setShowNameManager] = useState(false)
  const [showPivotDialog, setShowPivotDialog] = useState(false)
  const [pivotPane, setPivotPane] = useState<PivotPaneState | null>(null)
  /// Pane the user closed for this pivot; reopens once the selection leaves it.
  const dismissedPivotRef = useRef<string | null>(null)
  const openPivotPane = (): void => {
    const seed = onGetPivotEditSeed()
    if (!seed) return
    dismissedPivotRef.current = null
    setPivotPane({ seed, model: seed.model, deferred: false, dirty: false, error: null })
  }
  const applyPivotModel = (pane: PivotPaneState, model: PivotLayoutModel): PivotPaneState => {
    const invalid = validatePivotModel(model)
    if (invalid !== null)
      return { ...pane, model, dirty: false, error: t(PIVOT_MODEL_ERRORS[invalid]) }
    const failure = onEditPivot(modelToConfig(model, pane.seed.sourceRange, pane.seed.targetCell))
    return { ...pane, model, dirty: false, error: failure }
  }
  useEffect(() => {
    if (pivotSelectionKey === null) {
      // Leaving the pivot hides the pane unless deferred edits would be lost,
      // and a dismissed pane may open again on the next visit.
      dismissedPivotRef.current = null
      setPivotPane((current) => (current && !(current.deferred && current.dirty) ? null : current))
      return
    }
    if (pivotPane?.seed.key === pivotSelectionKey) return
    // Another pivot: the pane always follows the cursor so edits never land
    // on the previous table; a missing seed leaves it closed.
    setPivotPane(null)
    if (dismissedPivotRef.current !== pivotSelectionKey) openPivotPane()
    // Only the pivot identity drives this; the callbacks are App-scoped closures.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pivotSelectionKey])
  /** null = closed; otherwise the catalog category to open on and the cell the formula lands in */
  const [insertFunction, setInsertFunction] = useState<{
    category: string
    target: FunctionTarget
  } | null>(null)
  const insertFunctionOpen = insertFunction !== null
  const liveFunctions = useMemo(
    () => (insertFunctionOpen ? onListFunctions() : []),
    [insertFunctionOpen],
  )
  const insertFunctionOpenRef = useRef(insertFunctionOpen)
  insertFunctionOpenRef.current = insertFunctionOpen
  const openInsertFunction = (category: string): void => {
    // The modeless Function Arguments card keeps its cell: a grid pick moved
    // the active cell, so a second fx click must not re-target the session.
    if (insertFunctionOpenRef.current) return
    void functionArguments.readTarget().then((target) => {
      if (target && !insertFunctionOpenRef.current) setInsertFunction({ category, target })
    })
  }
  const [showSubtotalDialog, setShowSubtotalDialog] = useState(false)
  const [showGoalSeek, setShowGoalSeek] = useState(false)
  const [showFillSeries, setShowFillSeries] = useState(false)
  const [showConsolidateDialog, setShowConsolidateDialog] = useState(false)
  const [showGoTo, setShowGoTo] = useState(false)
  const [pasteSpecialSource, setPasteSpecialSource] = useState<PasteSpecialSource | null>(null)
  const [showZoomDialog, setShowZoomDialog] = useState(false)
  const [statsMenuAt, setStatsMenuAt] = useState<{ x: number; y: number } | null>(null)
  const openStatsMenu = (event: { clientX: number; clientY: number; preventDefault(): void }) => {
    event.preventDefault()
    setStatsMenuAt({ x: event.clientX, y: event.clientY })
  }
  useEffect(() => {
    // Univer owns the footer that hosts the sheet tabs and the statistics;
    // right-clicking anywhere in it outside the tab strip opens our menu.
    // Everything else the decision covers is sheet-context-menu.ts's: the
    // grid right-click must cancel the DOM event or Electron's native menu
    // stacks on top of Univer's own (genoffice#1816).
    const onContextMenu = (event: MouseEvent): void => {
      const action = sheetContextMenuAction(event.target)
      if (action.statsMenu) openStatsMenu(event)
      else if (action.suppressNativeMenu) event.preventDefault()
    }
    document.addEventListener('contextmenu', onContextMenu)
    return () => document.removeEventListener('contextmenu', onContextMenu)
  }, [])
  const [showHeaderFooter, setShowHeaderFooter] = useState(false)
  const [printRevision, setPrintRevision] = useState(0)
  const [showAllowEditRanges, setShowAllowEditRanges] = useState(false)
  const [protectSheetMode, setProtectSheetMode] = useState<'protect' | 'unprotect' | null>(null)
  /// Non-null while the Chart Design → Add Chart Element text prompt is open.
  const [chartTextTarget, setChartTextTarget] = useState<ChartTextTarget | null>(null)
  const onCommandRef = useRef(onCommand)
  const onIsCellEditingRef = useRef(onIsCellEditing)
  const openingWorkbookRef = useRef(openingWorkbook)
  onCommandRef.current = onCommand
  onIsCellEditingRef.current = onIsCellEditing
  openingWorkbookRef.current = openingWorkbook
  const dispatchCommand = (command: string): void => {
    if (command === 'format-cells') setShowFormatCells(true)
    else if (command === 'row-height-open') setAxisSizeTarget('row')
    else if (command === 'col-width-open') setAxisSizeTarget('col')
    else if (command === 'link-open') setShowLinkDialog(true)
    else if (command === 'sort-custom-open') setShowSortDialog(true)
    else if (command === 'remove-duplicates-open') setShowDedupeDialog(true)
    else if (command === 'table-create-open') setShowCreateTable(true)
    else if (command === 'table-resize-open') setShowResizeTable(true)
    else if (command === 'name-manager-open') setShowNameManager(true)
    else if (command === 'pivot-open') setShowPivotDialog(true)
    else if (command === 'pivot-edit') openPivotPane()
    else if (command === 'insert-function-open') openInsertFunction('All')
    else if (command.startsWith('insert-function-open:'))
      openInsertFunction(command.slice('insert-function-open:'.length))
    else if (command === 'goal-seek-open') setShowGoalSeek(true)
    else if (command === 'fill-series-open') setShowFillSeries(true)
    else if (command === 'subtotal-open') setShowSubtotalDialog(true)
    else if (command === 'consolidate-open') setShowConsolidateDialog(true)
    else if (command === 'goto-open') setShowGoTo(true)
    else if (command === PASTE_SPECIAL_OPEN_SHELL_COMMAND) {
      void onReadPasteSpecialSource().then((source) => setPasteSpecialSource(source))
    } else if (command === 'header-footer-open') setShowHeaderFooter(true)
    else if (command === 'allow-edit-ranges-open') setShowAllowEditRanges(true)
    else if (command === 'sheet-protect') {
      const protection = onGetSheetProtection()
      if (!protection) onCommand('sheet-protect')
      else if (!protection.protected) setProtectSheetMode('protect')
      else if (protection.hasPassword) setProtectSheetMode('unprotect')
      else onCommand('sheet-unprotect')
    } else if (command === 'ai-open-panel') setIsCopilotOpen(true)
    else if (command === 'ai-toggle-panel') setIsCopilotOpen((v) => !v)
    else if (command === 'chart-element-title') setChartTextTarget('title')
    else if (command === 'chart-element-axis-cat') setChartTextTarget('axis-category')
    else if (command === 'chart-element-axis-val') setChartTextTarget('axis-value')
    else if (command === 'formula-bar-toggle') toggleFormulaBarExpand()
    else onCommand(command)
  }
  const dispatchCommandRef = useRef(dispatchCommand)
  dispatchCommandRef.current = dispatchCommand
  useEffect(() => {
    setContextMenuCommandSink((command) => dispatchCommandRef.current(command))
    const unregisterFx = registerFxButtonOpener(() =>
      dispatchCommandRef.current('insert-function-open'),
    )
    const unregisterPasteSpecial = registerPasteSpecialOpener(() =>
      dispatchCommandRef.current(PASTE_SPECIAL_OPEN_SHELL_COMMAND),
    )
    return () => {
      setContextMenuCommandSink(null)
      unregisterPasteSpecial()
      unregisterFx()
    }
  }, [])
  useEffect(() => {
    // Table + gating live in global-shortcuts.ts; text fields (AI chat,
    // dialogs, Univer's find/replace and formula bar) never trigger these.
    const onKeyDown = (event: KeyboardEvent): void => {
      if (openingWorkbookRef.current) return
      const action = resolveGlobalShortcut(event, {
        isMac: IS_MAC,
        modalOpen: isModalOpen(),
        cellEditing: onIsCellEditingRef.current(),
        gridTarget: isGridKeyTarget(event.target),
        formulaBarTarget: isFormulaBarKeyTarget(event.target),
      })
      if (!action) return
      event.preventDefault()
      if (action.kind === 'dialog') {
        if (action.dialog === 'formatCells') setShowFormatCells(true)
        else setShowGoTo(true)
        return
      }
      dispatchCommandRef.current(action.command)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])
  // Univer's shortcut dispatcher captures keydown and binds Backspace to
  // "delete-and-start-editing" (active cell only). Register on window
  // capture *here* (child effect runs before App creates Univer) so we
  // win the race and clear the whole selection instead. Keep the listener
  // mounted once: re-binding after Univer starts would lose capture order.
  useEffect(() => {
    const onKeyDownCapture = (event: KeyboardEvent): void => {
      if (openingWorkbookRef.current) return
      if (!shouldInterceptClearSelection(event, onIsCellEditingRef.current())) return
      event.preventDefault()
      event.stopImmediatePropagation()
      onCommandRef.current('clear-contents')
    }
    window.addEventListener('keydown', onKeyDownCapture, true)
    return () => window.removeEventListener('keydown', onKeyDownCapture, true)
  }, [])
  // Deselecting while on the contextual tab lands back on Home.
  useEffect(() => {
    if (!selectedChart && activeTab === 'Chart Design') setActiveTab('Home')
  }, [selectedChart, activeTab])
  // Leaving a table while on its tab lands back on Home.
  useEffect(() => {
    if (!selectedTable && activeTab === 'Table Design') setActiveTab('Home')
  }, [selectedTable, activeTab])
  useEffect(() => {
    if (
      (activeTab === 'Shape Format' && selectedVisualKind !== 'shape') ||
      (activeTab === 'Picture Format' && selectedVisualKind !== 'image')
    ) {
      setActiveTab('Home')
    }
  }, [selectedVisualKind, activeTab])
  // A freshly created table activates its tab (Excel). The create may finish
  // asynchronously (streamed rows load first), so the signal is a counter,
  // consumed only once the selection echo has caught up with the new table.
  const seenTableCreatedSeq = useRef(tableCreatedSeq)
  useEffect(() => {
    if (tableCreatedSeq === seenTableCreatedSeq.current || !selectedTable) return
    seenTableCreatedSeq.current = tableCreatedSeq
    setActiveTab('Table Design')
  }, [tableCreatedSeq, selectedTable])
  // Univer's formula-bar Name Box (the defined-name selector) is the only
  // cell-reference box; the Go To ▾ arrow rides inside the bar next to it.
  // The bar is Univer-owned DOM that can remount with the workbench, so the
  // button is (re)inserted on mutation rather than rendered by React.
  const gotoTip = t('appGoToButtonTitle')
  const gotoTipRef = useRef(gotoTip)
  gotoTipRef.current = gotoTip
  useEffect(() => {
    const button = document.querySelector<HTMLButtonElement>('.goto-in-bar')
    if (button) button.dataset['tip'] = gotoTip
  }, [gotoTip])
  useEffect(() => {
    const ensure = (): void => {
      // The Name Box sits in a fixed-width block wrapper; the flex row is the
      // bar itself, so the button must ride as the wrapper's sibling.
      const wrapper = document.querySelector('[data-u-comp="defined-name"]')?.parentElement
      const bar = wrapper?.parentElement
      if (!wrapper || !bar) return
      let button = bar.querySelector<HTMLButtonElement>('.goto-in-bar')
      if (!button) {
        button = document.createElement('button')
        button.type = 'button'
        button.className = 'goto-in-bar'
        button.textContent = '▾'
        button.setAttribute('aria-label', 'Go To')
        button.addEventListener('click', () => setShowGoTo(true))
        wrapper.after(button)
      }
      if (button.dataset['tip'] !== gotoTipRef.current) button.dataset['tip'] = gotoTipRef.current
    }
    ensure()
    const container = document.getElementById('univer-container')
    if (!container) return undefined
    const observer = new MutationObserver(ensure)
    observer.observe(container, { childList: true, subtree: true })
    return () => {
      observer.disconnect()
      document.querySelector('.goto-in-bar')?.remove()
    }
  }, [])
  const visibleTabs: readonly RibbonTab[] = [
    ...ribbonTabs,
    ...(selectedChart ? (['Chart Design'] as const) : []),
    ...(selectedTable ? (['Table Design'] as const) : []),
    ...(selectedVisualKind === 'shape' ? (['Shape Format'] as const) : []),
    ...(selectedVisualKind === 'image' ? (['Picture Format'] as const) : []),
  ]
  const saveAsTitle = `${t('appSaveAs')} (${platformShortcuts('⇧⌘S')})`

  return (
    <main
      className={`app-shell ${isCopilotOpen ? '' : 'copilot-collapsed'}`}
      inert={openingWorkbook}
      aria-busy={openingWorkbook}
    >
      <header className={`excel-header ${collapse.rootClass}`} ref={collapse.rootRef}>
        <nav
          className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}`}
          aria-label={t('appScopeWorkbook')}
          onDoubleClick={collapse.onTabsDoubleClick}
        >
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appSaveTitle')}
            aria-label={t('appSaveTitle')}
            disabled={!canSave}
            onClick={onSave}
          >
            <SaveIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={saveAsTitle}
            aria-label={saveAsTitle}
            disabled={!canSaveAs}
            onClick={onSaveAs}
          >
            <SaveAsIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appUndo')}
            aria-label={t('appUndo')}
            disabled={!canUndo}
            onClick={() => onUndo()}
          >
            <UndoIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appRedo')}
            aria-label={t('appRedo')}
            disabled={!canRedo}
            onClick={onRedo}
          >
            <RedoIcon />
          </button>
          <label
            className={`autosave-toggle ${autoSave ? 'on' : ''}`}
            data-tip={t('appAutoSaveTip')}
          >
            <span className="autosave-knob" />
            <span className="autosave-text">{t('appAutoSave')}</span>
            <input
              type="checkbox"
              checked={autoSave}
              onChange={(e) => onAutoSaveChange(e.target.checked)}
            />
          </label>
          <span className="qa-sep" aria-hidden="true" />
          {visibleTabs.map((tab) => (
            <button
              className={`${collapse.tabClass(tab === activeTab)} ${CONTEXTUAL_TABS.includes(tab) ? 'contextual' : ''}`}
              data-tip={collapse.tabTip(tab === activeTab)}
              key={tab}
              onClick={() => {
                collapse.onTabPress(tab === activeTab)
                setActiveTab(tab)
              }}
            >
              {t(TAB_LABEL[tab])}
            </button>
          ))}
          <span className="ribbon-tabs-spacer" />
          <span className="workbook-status" role="status" aria-live="polite">
            {statusMessage}
          </span>
        </nav>

        <Ribbon
          activeTab={activeTab}
          selectionFormat={selectionFormat}
          formatPainterActive={formatPainterActive}
          sheetHasContent={sheetHasContent}
          sheetProtected={onGetSheetProtection()?.protected ?? null}
          workbookProtected={onGetWorkbookProtection()}
          formulaBarVisible={formulaBarVisible}
          crossHighlightVisible={crossHighlightVisible}
          pageLayout={pageLayout}
          selectedChart={selectedChart}
          selectedShape={selectedShape}
          selectedTable={selectedTable}
          outlineSummary={outlineSummary}
          onListNames={() => {
            // Names scoped to another sheet resolve to #NAME? here; only
            // workbook-scoped and active-sheet names are usable in a formula.
            const data = onGetDefinedNames()
            return data.names
              .filter(
                (entry) => entry.scopeSheetId === null || entry.scopeSheetId === data.activeSheetId,
              )
              .map((entry) => entry.name)
          }}
          calcManual={calcManual}
          onRefreshPivot={onRefreshPivot}
          onIsSelectionInPivot={onIsSelectionInPivot}
          onCommand={dispatchCommand}
          onAiRun={(nextPrompt) => {
            setIsCopilotOpen(true)
            onSend(nextPrompt)
          }}
          aiOpen={isCopilotOpen}
          onAiToggle={() => setIsCopilotOpen((open) => !open)}
        />
      </header>

      {/* AI panel docks on the left, full height under the ribbon (unified with docs) */}
      <div className="sheet-body">
        <AiChatPanel
          isOpen={isCopilotOpen}
          hasContent={sheetHasContent}
          chat={chat}
          {...(historicChat !== undefined ? { historicChat } : {})}
          attachments={attachments}
          attachNotice={attachNotice}
          onPickAttachments={onPickAttachments}
          onAddAttachmentPaths={onAddAttachmentPaths}
          onAddPastedImage={onAddPastedImage}
          onRemoveAttachment={onRemoveAttachment}
          prompt={prompt}
          preview={preview}
          aiBusy={aiBusy}
          onPromptChange={onPromptChange}
          onSend={onSend}
          onSendQueued={(text, atts) => {
            if (aiBusy) return false
            onSend(text, atts)
            return true
          }}
          onStop={onStop}
          onNewChat={onNewChat}
          onUndo={onUndo}
          scopeRange={aiScopeRange}
          scopeColumns={aiScopeColumns}
          scopeLocked={aiScopeLocked}
          onScopeDismiss={onAiScopeDismiss}
          onCitation={onAiCitation}
          onExpand={() => setIsCopilotOpen(true)}
          onCollapse={() => setIsCopilotOpen(false)}
        />
        <div className="sheet-main">
          <section className="workbook-area">
            <div id="univer-container" className="spreadsheet" />
            {emptyCsvNotice && (
              <div className="csv-empty-notice" role="status">
                <span>{t('appCsvEmpty')}</span>
                <button type="button" onClick={onOpenWorkbook}>
                  {t('appOpenWorkbookTitle')}
                </button>
                <button
                  type="button"
                  className="csv-empty-notice-close"
                  aria-label={t('appClose')}
                  onClick={onDismissEmptyCsvNotice}
                >
                  ×
                </button>
              </div>
            )}
          </section>
          {aiSelectionAskAnchor && aiScopeRange && !aiBusy && (
            <AiSelectionAsk
              anchor={aiSelectionAskAnchor}
              range={aiScopeRange}
              onDismiss={onAiSelectionAskDismiss}
              onSend={(instruction) => {
                setIsCopilotOpen(true)
                onSend(instruction)
              }}
            />
          )}

          {/* Status bar spans the sheet column only — the AI dock keeps the full window height (unified with docs/slides). */}
          <footer className="status-bar" onContextMenu={openStatsMenu}>
            <div className="status-left">
              <span className="status-msg">{statusMessage}</span>
            </div>
            <div className="status-right">
              {/* One-of-N view switcher, so it is a radiogroup rather than a
                  second set of buttons named like the View tab's. */}
              <div
                className="view-switch"
                role="radiogroup"
                aria-label={t('appGroupWorkbookViews')}
              >
                <button
                  className="view-btn"
                  role="radio"
                  data-tip={t('appNormalViewTip')}
                  aria-label={t('appNormalViewTip')}
                  aria-checked={pageLayout.pageBreakPreview !== true}
                  onClick={() => {
                    if (pageLayout.pageBreakPreview === true) onCommand('toggle-page-break-preview')
                  }}
                >
                  ▦
                </button>
                <button
                  className="view-btn"
                  role="radio"
                  data-tip={t('appPageBreakPreviewTip')}
                  aria-label={t('appPageBreakPreviewTip')}
                  aria-checked={pageLayout.pageBreakPreview === true}
                  onClick={() => {
                    if (pageLayout.pageBreakPreview !== true) onCommand('toggle-page-break-preview')
                  }}
                >
                  ┆
                </button>
              </div>
              <button
                className="zoom-btn"
                data-tip={t('appZoomOut')}
                aria-label={t('appZoomOut')}
                onClick={() => onCommand('zoom-out')}
              >
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={SHEET_ZOOM_MIN}
                max={SHEET_ZOOM_MAX}
                value={clampZoomPercent(zoomPercent)}
                onChange={(event) => onCommand(`zoom:${event.target.value}`)}
              />
              <button
                className="zoom-btn"
                data-tip={t('appZoomIn')}
                aria-label={t('appZoomIn')}
                onClick={() => onCommand('zoom-in')}
              >
                +
              </button>
              <button
                className="zoom-value"
                data-tip={t('appZoomLevel')}
                aria-label={t('appZoomLevel')}
                onClick={() => setShowZoomDialog(true)}
              >
                {zoomPercent}%
              </button>
            </div>
          </footer>
          {statsMenuAt && (
            <StatusBarStatsMenu
              at={statsMenuAt}
              enabled={statusBarFuncs}
              onToggle={onToggleStatusBarFunc}
              onClose={() => setStatsMenuAt(null)}
            />
          )}
        </div>
      </div>
      {showFormatCells && (
        <FormatCellsDialog
          selectionFormat={selectionFormat}
          anchorValue={onGetAnchorValue()}
          onCommand={onCommand}
          onClose={() => setShowFormatCells(false)}
        />
      )}
      {axisSizeTarget && (
        <AxisSizeDialog
          axis={axisSizeTarget}
          onCommand={onCommand}
          onClose={() => setAxisSizeTarget(null)}
        />
      )}
      {showLinkDialog && (
        <LinkDialog
          currentTarget={selectionFormat?.link ?? null}
          onCommand={onCommand}
          onClose={() => setShowLinkDialog(false)}
        />
      )}
      {chartTextTarget && selectedChart && (
        <ChartTextDialog
          target={chartTextTarget}
          initial={
            chartTextTarget === 'title'
              ? selectedChart.title
              : ((chartTextTarget === 'axis-category'
                  ? selectedChart.axisTitles?.category
                  : selectedChart.axisTitles?.value) ?? '')
          }
          onCommand={onCommand}
          onClose={() => setChartTextTarget(null)}
        />
      )}
      {showSortDialog && (
        <SortDialog
          columns={onGetSortColumns()}
          onCommand={onCommand}
          onClose={() => setShowSortDialog(false)}
        />
      )}
      {showDedupeDialog && (
        <RemoveDuplicatesDialog onCommand={onCommand} onClose={() => setShowDedupeDialog(false)} />
      )}
      {showCreateTable && (
        <TableRangeDialog
          title={t('appCreateTable')}
          note={t('appCreateTableWhere')}
          initialRange={onGetCreateTableRange()}
          headersOption
          onSubmit={(range, headers) => dispatchCommand(`table-create:${range}:${headers ? 1 : 0}`)}
          onClose={() => setShowCreateTable(false)}
        />
      )}
      {showResizeTable && selectedTable && (
        <TableRangeDialog
          title={t('appResizeTable')}
          note={t('appResizeTableNote')}
          initialRange={selectedTable.rangeA1}
          onSubmit={(range) => onCommand(`table-resize:${range}`)}
          onClose={() => setShowResizeTable(false)}
        />
      )}
      {showNameManager &&
        (() => {
          const data = onGetDefinedNames()
          return (
            <NameManagerDialog
              names={data.names}
              sheets={data.sheets}
              onAction={onDefinedNameAction}
              onClose={() => setShowNameManager(false)}
            />
          )
        })()}
      {showPivotDialog &&
        (() => {
          const sourceRange = onGetSourceRange()
          return (
            <PivotDialog
              hasFields={onGetPivotFields(sourceRange).length > 0}
              sourceRange={sourceRange}
              onCreate={(request) => {
                const failure = onCreatePivot(request)
                if (failure === null) openPivotPane()
                return failure
              }}
              onClose={() => setShowPivotDialog(false)}
            />
          )
        })()}
      {pivotPane && (
        <PivotFieldPane
          fields={pivotPane.seed.fields}
          model={pivotPane.model}
          deferred={pivotPane.deferred}
          dirty={pivotPane.dirty}
          error={pivotPane.error}
          onModelChange={(model) =>
            setPivotPane(
              pivotPane.deferred
                ? { ...pivotPane, model, dirty: true, error: null }
                : applyPivotModel(pivotPane, model),
            )
          }
          onDeferredChange={(deferred) =>
            setPivotPane(
              !deferred && pivotPane.dirty
                ? { ...applyPivotModel(pivotPane, pivotPane.model), deferred }
                : { ...pivotPane, deferred },
            )
          }
          onUpdate={() => setPivotPane(applyPivotModel(pivotPane, pivotPane.model))}
          onGetFieldMembers={onGetPivotFieldMembers}
          onClose={() => {
            dismissedPivotRef.current = pivotPane.seed.key
            setPivotPane(null)
          }}
        />
      )}
      {showFillSeries && (
        <FillSeriesDialog
          context={onGetFillSeriesContext()}
          defaultType={fillSeriesDefaultType(selectionFormat?.numberFormat)}
          onApply={onFillSeries}
          onClose={() => setShowFillSeries(false)}
        />
      )}
      {showGoalSeek && (
        <GoalSeekDialog
          initialSetCell={onGetActiveCell()}
          onSolve={onGoalSeek}
          onClose={() => setShowGoalSeek(false)}
        />
      )}
      {insertFunction !== null && (
        <InsertFunctionFlow
          functions={liveFunctions}
          target={insertFunction.target}
          initialCategory={insertFunction.category}
          host={functionArguments}
          onClose={() => setInsertFunction(null)}
        />
      )}
      {showSubtotalDialog && (
        <SubtotalDialog
          fields={onGetPivotFields()}
          onCreate={onCreateSubtotal}
          onClose={() => setShowSubtotalDialog(false)}
        />
      )}
      {showConsolidateDialog && (
        <ConsolidateDialog
          defaultReference={onGetConsolidateDefault()}
          targetLabel={onGetActiveCell()}
          onCreate={onCreateConsolidate}
          onClose={() => setShowConsolidateDialog(false)}
        />
      )}
      {showZoomDialog && (
        <ZoomDialog
          zoomPercent={zoomPercent}
          onCommand={onCommand}
          onClose={() => setShowZoomDialog(false)}
        />
      )}
      {pasteSpecialSource && (
        <PasteSpecialDialog
          source={pasteSpecialSource}
          onApply={(options) => onPasteSpecial(pasteSpecialSource, options)}
          onClose={() => setPasteSpecialSource(null)}
        />
      )}
      {showGoTo && (
        <GoToDialog
          names={onListDefinedNames()}
          onGo={onGoToReference}
          onClose={() => setShowGoTo(false)}
        />
      )}
      {protectSheetMode && (
        <ProtectSheetDialog
          mode={protectSheetMode}
          onProtect={onProtectSheet}
          onUnprotect={onUnprotectSheet}
          onClose={() => setProtectSheetMode(null)}
        />
      )}
      {showAllowEditRanges &&
        (() => {
          const snapshot = onGetProtectedRanges()
          return (
            <AllowEditRangesDialog
              ranges={snapshot.error === null ? snapshot.ranges : []}
              defaultRef={activeCellA1}
              onApply={(ranges) => snapshot.error ?? onApplyProtectedRanges(ranges)}
              onClose={() => setShowAllowEditRanges(false)}
            />
          )
        })()}
      {printHost && (
        <PrintDialog
          host={printHost}
          onClose={onClosePrint}
          onPageSetup={() => setShowHeaderFooter(true)}
          revision={printRevision}
        />
      )}
      {showHeaderFooter && (
        <HeaderFooterDialog
          initialHeader={pageLayout.header ?? null}
          initialFooter={pageLayout.footer ?? null}
          onApply={onApplyHeaderFooter}
          onClose={() => {
            setShowHeaderFooter(false)
            setPrintRevision((value) => value + 1)
          }}
        />
      )}
      <SheetTabContextMenu
        actions={sheetTabActions}
        getSheetProtected={() => onGetSheetProtection()?.protected ?? null}
        getWorkbookLocked={onGetWorkbookProtection}
        onProtectSheet={() => dispatchCommand('sheet-protect')}
      />
    </main>
  )
}

const NO_SORT_LEVEL = -1

function SortDialog({
  columns,
  onCommand,
  onClose,
}: {
  readonly columns: { label: string; colIndex: number }[]
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [hasHeader, setHasHeader] = useState(true)
  const first = columns[0]?.colIndex ?? NO_SORT_LEVEL
  const [levels, setLevels] = useState<{ colIndex: number; order: 'a' | 'd' }[]>([
    { colIndex: first, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
  ])
  const apply = (): void => {
    const rules = levels
      .filter((level) => level.colIndex !== NO_SORT_LEVEL)
      .map((level) => `${level.colIndex}${level.order}`)
    if (rules.length > 0) onCommand(`sort-custom:${hasHeader ? 1 : 0}:${rules.join(',')}`)
    onClose()
  }
  const setLevel = (
    index: number,
    patch: Partial<{ colIndex: number; order: 'a' | 'd' }>,
  ): void => {
    setLevels((previous) =>
      previous.map((level, at) => (at === index ? { ...level, ...patch } : level)),
    )
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appCustomSort')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appSort')}</header>
        <div className="dialog-body">
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
          {levels.map((level, index) => (
            <div className="sort-level" key={index}>
              <span>{t(index === 0 ? 'appSortBy' : 'appThenBy')}</span>
              <Dropdown
                ariaLabel={t(index === 0 ? 'appSortBy' : 'appThenBy')}
                value={String(level.colIndex)}
                options={[
                  ...(index > 0 ? [{ value: String(NO_SORT_LEVEL), label: t('appSortNone') }] : []),
                  ...columns.map((column) => ({
                    value: String(column.colIndex),
                    label: column.label,
                  })),
                ]}
                onPick={(v) => setLevel(index, { colIndex: Number(v) })}
              />
              <Dropdown
                value={level.order}
                options={[
                  { value: 'a', label: t('appSortAsc') },
                  { value: 'd', label: t('appSortDesc') },
                ]}
                onPick={(v) => setLevel(index, { order: v })}
              />
            </div>
          ))}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            disabled={levels.every((level) => level.colIndex === NO_SORT_LEVEL)}
            onClick={apply}
          >
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

function RemoveDuplicatesDialog({
  onCommand,
  onClose,
}: {
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [hasHeader, setHasHeader] = useState(true)
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appRemoveDuplicates')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appRemoveDuplicates')}</header>
        <div className="dialog-body">
          <p className="dialog-note">{t('appRemoveDuplicatesNote')}</p>
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            onClick={() => {
              onCommand(`remove-duplicates:${hasHeader ? 1 : 0}`)
              onClose()
            }}
          >
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Table Name box: commits on Enter / blur, Esc restores.
function TableNameInput({
  name,
  onCommit,
}: {
  readonly name: string
  readonly onCommit: (name: string) => void
}): React.JSX.Element {
  const [value, setValue] = useState(name)
  // Escape blurs the box; the blur must then restore, not commit the stale text.
  const cancelledRef = useRef(false)
  const commit = (): void => {
    if (cancelledRef.current) {
      cancelledRef.current = false
      setValue(name)
      return
    }
    const next = value.trim()
    if (next.length > 0 && next !== name) onCommit(next)
    else setValue(name)
  }
  return (
    <input
      className="table-name-input"
      value={value}
      spellCheck={false}
      onChange={(event) => setValue(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          commit()
          event.currentTarget.blur()
        } else if (event.key === 'Escape') {
          cancelledRef.current = true
          setValue(name)
          event.currentTarget.blur()
        }
        event.stopPropagation()
      }}
    />
  )
}

/// Excel's Table Styles gallery: Light / Medium / Dark rows of swatches
/// drawn from the same palette rules the grid renders with.
function TableStyleGallery({
  current,
  onPick,
}: {
  readonly current: string | null
  readonly onPick: (style: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  const families: { family: TableStyleFamily; label: string }[] = [
    { family: 'Light', label: t('appTableStyleFamilyLight') },
    { family: 'Medium', label: t('appTableStyleFamilyMedium') },
    { family: 'Dark', label: t('appTableStyleFamilyDark') },
  ]
  const swatch = (style: string, label: string): React.JSX.Element => {
    const palette = builtinTablePalette(style, undefined)
    const rowFill = (index: number): string | undefined =>
      index % 2 === 0 ? palette.stripeFill : (palette.secondRowStripeFill ?? palette.wholeTableFill)
    return (
      <button
        key={style}
        className={`table-style-swatch ${current === style ? 'active' : ''}`}
        data-tip={label}
        onClick={() => {
          onPick(style)
          setOpen(false)
        }}
      >
        <span
          className="swatch-row header"
          style={{
            background: palette.headerFill,
            borderColor: palette.borderColor ?? palette.wholeTableBorderColor ?? palette.headerFill,
          }}
        >
          {[0, 1, 2].map((cell) => (
            <i key={cell} style={{ background: palette.headerFontColor }} />
          ))}
        </span>
        {[0, 1, 2, 3].map((row) => (
          <span
            key={row}
            className="swatch-row"
            style={{ background: rowFill(row), borderColor: palette.innerHorizontalBorderColor }}
          >
            {[0, 1, 2].map((cell) => (
              <i key={cell} style={{ background: palette.bodyFontColor }} />
            ))}
          </span>
        ))}
      </button>
    )
  }
  const currentLabel = (() => {
    const match = current ? /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(current) : null
    if (!match) return t('appGroupTableStyles')
    const family = families.find((entry) => entry.family === match[1])
    return `${family?.label ?? match[1]} ${match[2]}`
  })()
  return (
    <div className="table-style-gallery" ref={wrapRef}>
      <button
        className="styles-row as-button"
        data-tip={t('appTableStylesGalleryTip')}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
      >
        <ToolSymbol symbol="▦" />
        {currentLabel}
        <CaretIcon />
      </button>
      {open && (
        <div className="table-style-panel" role="dialog" aria-label={t('appGroupTableStyles')}>
          {families.map(({ family, label }) => (
            <section key={family}>
              <header>{label}</header>
              <div className="table-style-grid">
                {tableStyleNames(family).map((style, index) =>
                  swatch(style, `${label} ${index + 1}`),
                )}
              </div>
            </section>
          ))}
        </div>
      )}
    </div>
  )
}

/// Create Table (Ctrl+T) / Resize Table: one range box, optionally the
/// "My table has headers" check.
function TableRangeDialog({
  title,
  note,
  initialRange,
  headersOption = false,
  onSubmit,
  onClose,
}: {
  readonly title: string
  readonly note: string
  readonly initialRange: string
  readonly headersOption?: boolean
  readonly onSubmit: (range: string, headers: boolean) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [range, setRange] = useState(initialRange)
  const [headers, setHeaders] = useState(true)
  const modal = useModalDialog(onClose)
  const submit = (): void => {
    if (range.trim().length === 0) return
    onSubmit(range.trim().toUpperCase(), headers)
    onClose()
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{title}</header>
        <div className="dialog-body">
          <p className="dialog-note">{note}</p>
          <input
            className="table-range-input"
            autoFocus
            value={range}
            spellCheck={false}
            onChange={(event) => setRange(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') submit()
            }}
          />
          {headersOption && (
            <label className="sort-header-check">
              <input
                type="checkbox"
                checked={headers}
                onChange={(event) => setHeaders(event.target.checked)}
              />
              {t('appMyTableHasHeaders')}
            </label>
          )}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" onClick={submit}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Chart Design → Add Chart Element text prompt (title / axis titles).
/// An empty axis title removes that axis title; an empty chart title blanks it.
function ChartTextDialog({
  target,
  initial,
  onCommand,
  onClose,
}: {
  readonly target: ChartTextTarget
  readonly initial: string
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(initial)
  const { heading, command } = CHART_TEXT_LABELS[target]
  const apply = (): void => {
    onCommand(`${command}:${value.trim()}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(heading)}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(heading)}</header>
        <div className="dialog-body">
          <label>
            {t(target === 'title' ? 'appTitleText' : 'appAxisTitleText')}
            <input
              autoFocus
              type="text"
              value={value}
              maxLength={255}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Numeric row-height / column-width entry (Excel's Format → Row Height /
/// Column Width). Applies to the rows/columns of the current selection.
function AxisSizeDialog({
  axis,
  onCommand,
  onClose,
}: {
  readonly axis: 'row' | 'col'
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState('')
  const max = axis === 'row' ? 409.5 : 255
  const parsed = Number(value.trim().replace(',', '.'))
  const valid = value.trim() !== '' && Number.isFinite(parsed) && parsed >= 0 && parsed <= max
  const apply = (): void => {
    if (!valid) return
    onCommand(`${axis === 'row' ? 'row-height' : 'col-width'}:${parsed}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}</header>
        <div className="dialog-body">
          <label>
            {t(axis === 'row' ? 'appRowHeightLabel' : 'appColWidthLabel')}
            <input
              autoFocus
              type="text"
              inputMode="decimal"
              value={value}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" disabled={!valid} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

function LinkDialog({
  currentTarget,
  onCommand,
  onClose,
}: {
  readonly currentTarget: string | null
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(currentTarget ?? '')
  const apply = (): void => {
    if (value.trim()) onCommand(`link-set:${encodeURIComponent(value)}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}</header>
        <div className="dialog-body">
          <label>
            {t('appLinkAddressLabel')}
            <input
              autoFocus
              type="text"
              value={value}
              placeholder={t('appLinkPlaceholder')}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          {currentTarget && (
            <button
              onClick={() => {
                onCommand('link-remove')
                onClose()
              }}
            >
              {t('appRemoveLink')}
            </button>
          )}
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" disabled={!value.trim()} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

function Ribbon({
  activeTab,
  selectionFormat,
  formatPainterActive,
  sheetHasContent,
  sheetProtected,
  workbookProtected,
  formulaBarVisible,
  crossHighlightVisible,
  pageLayout,
  selectedChart,
  selectedShape,
  selectedTable,
  outlineSummary,
  onCommand,
  onAiRun,
  aiOpen,
  onAiToggle,
  onListNames,
  calcManual,
  onRefreshPivot,
  onIsSelectionInPivot,
}: {
  readonly activeTab: RibbonTab
  readonly selectionFormat: SelectionFormat | null
  readonly formatPainterActive: boolean
  readonly sheetHasContent: boolean
  readonly sheetProtected: boolean | null
  readonly workbookProtected: boolean | null
  /// View > Show echo for the formula bar toggle (app-level, not per sheet).
  readonly formulaBarVisible: boolean
  readonly crossHighlightVisible: boolean
  readonly pageLayout: PageLayoutEcho
  readonly selectedChart: SelectedChartRibbon | null
  readonly selectedShape: SelectedShapeRibbon | null
  readonly selectedTable: SelectedTableRibbon | null
  readonly outlineSummary: { below: boolean; right: boolean }
  readonly onCommand: (command: string) => void
  /** Defined names for the Use in Formula menu. */
  readonly onListNames: () => readonly string[]
  /** Manual-recalc mode echo for the Calculation Options menu. */
  readonly calcManual: boolean
  /** Open the AI panel and immediately send the given prompt */
  readonly onAiRun: (prompt: string) => void
  /** AI side panel visibility (docs/slides parity: the entry button toggles it) */
  readonly aiOpen: boolean
  readonly onAiToggle: () => void
  readonly onRefreshPivot: () => string | null
  readonly onIsSelectionInPivot: () => boolean
}): React.JSX.Element {
  const { t } = useI18n()
  const [fontColor, setFontColor] = useState('#C00000')
  const [fillColor, setFillColor] = useState('#FFF2CC')
  const [borderColor, setBorderColor] = useState('#000000')
  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // Large menu button: a native select stretched invisibly over the tool,
  // each option carrying its full command string.
  const largeMenu = (
    label: string,
    symbol: string,
    title: string,
    options: readonly { value: string; label: string }[],
  ): React.JSX.Element => (
    <div className="ribbon-tool large" data-tip={title}>
      <span className="tool-icon-row">
        <ToolSymbol symbol={symbol} />
        <CaretIcon />
      </span>
      <span>
        <strong>{label}</strong>
      </span>
      <MenuSelect cover label={label} options={options} onPick={onCommand} />
    </div>
  )
  if (activeTab === 'Chart Design') {
    const canConvert = Boolean(selectedChart?.convertible && selectedChart.canEdit)
    const canEditChart = Boolean(selectedChart?.canEdit)
    const labelValue = selectedChart?.isPie ? 'percent' : 'value'
    // ✓ mirrors the chart's current state (parsed from the file + pending
    // edits), so the menu reads as checked toggles.
    const activeLabels = selectedChart?.dataLabels ?? 'none'
    const activeLegend = selectedChart?.legend ?? 'none'
    const check = (on: boolean): string => (on ? ' ✓' : '')
    const elementOptions = [
      { value: 'chart-element-title', label: t('appChartElTitle') },
      ...(selectedChart?.hasAxes
        ? [
            {
              value: 'chart-element-axis-cat',
              label: t('appChartElAxisCat') + check(Boolean(selectedChart.axisTitles?.category)),
            },
            {
              value: 'chart-element-axis-val',
              label: t('appChartElAxisVal') + check(Boolean(selectedChart.axisTitles?.value)),
            },
          ]
        : []),
      ...(selectedChart?.canLabel
        ? [
            ...(selectedChart.isPie
              ? [
                  {
                    value: 'chart-labels:category-percent',
                    label:
                      t('appChartElLabelsNamePct') + check(activeLabels === 'category-percent'),
                  },
                ]
              : []),
            {
              value: `chart-labels:${labelValue}`,
              label:
                t(selectedChart.isPie ? 'appChartElLabelsPct' : 'appChartElLabelsValues') +
                check(activeLabels === labelValue),
            },
            {
              value: 'chart-labels:none',
              label: t('appChartElLabelsNone') + check(activeLabels === 'none'),
            },
          ]
        : []),
      {
        value: 'chart-legend:right',
        label: t('appChartElLegendRight') + check(activeLegend === 'right'),
      },
      {
        value: 'chart-legend:top',
        label: t('appChartElLegendTop') + check(activeLegend === 'top'),
      },
      {
        value: 'chart-legend:bottom',
        label: t('appChartElLegendBottom') + check(activeLegend === 'bottom'),
      },
      {
        value: 'chart-legend:left',
        label: t('appChartElLegendLeft') + check(activeLegend === 'left'),
      },
      {
        value: 'chart-legend:none',
        label: t('appChartElLegendNone') + check(activeLegend === 'none'),
      },
    ]
    const layoutLabels = t(selectedChart?.isPie ? 'appLayoutLabelsNamePct' : 'appLayoutLabelsValue')
    const layoutOptions = [
      { value: 'chart-layout:1', label: t('appLayout1', { labels: layoutLabels }) },
      { value: 'chart-layout:2', label: t('appLayout2', { labels: layoutLabels }) },
      { value: 'chart-layout:3', label: t('appLayout3') },
      { value: 'chart-layout:4', label: t('appLayout4', { labels: layoutLabels }) },
    ]
    const colorOptions = [
      { value: 'chart-colors:office', label: t('appColorsOffice') },
      { value: 'chart-colors:blue', label: t('appColorsBlue') },
      { value: 'chart-colors:green', label: t('appColorsGreen') },
      { value: 'chart-colors:warm', label: t('appColorsWarm') },
      { value: 'chart-colors:gray', label: t('appColorsGray') },
    ]
    const canRecolor =
      canEditChart &&
      ((selectedChart?.isPie ? selectedChart.categoryCount : selectedChart?.seriesCount) ?? 0) > 0
    const stackable =
      canEditChart && ['column', 'bar', 'line', 'area'].includes(selectedChart?.currentType ?? '')
    // 'standard' is how line/area spell side-by-side in OOXML.
    const activeGrouping =
      selectedChart?.grouping === 'stacked' || selectedChart?.grouping === 'percentStacked'
        ? selectedChart.grouping
        : 'clustered'
    const groupingOptions = [
      {
        value: 'chart-grouping:clustered',
        label: t('appGroupingClustered') + check(activeGrouping === 'clustered'),
      },
      {
        value: 'chart-grouping:stacked',
        label: t('appGroupingStacked') + check(activeGrouping === 'stacked'),
      },
      {
        value: 'chart-grouping:percentStacked',
        label: t('appGroupingPercent') + check(activeGrouping === 'percentStacked'),
      },
    ]
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupChartLayouts')}>
          {canEditChart ? (
            largeMenu(t('appAddChartElement'), '📊', t('appAddChartElementTitle'), elementOptions)
          ) : (
            <RibbonButton
              large
              menu
              label={t('appAddChartElement')}
              detail={t('appSelectEditableChart')}
              symbol="📊"
              disabled
              onClick={() => {}}
            />
          )}
          {canEditChart ? (
            largeMenu(t('appQuickLayout'), '▦', t('appQuickLayoutTitle'), layoutOptions)
          ) : (
            <RibbonButton
              large
              menu
              label={t('appQuickLayout')}
              detail={t('appSelectEditableChart')}
              symbol="▦"
              disabled
              onClick={() => {}}
            />
          )}
          {canRecolor ? (
            largeMenu(
              t('appChangeColors'),
              '🎨',
              t(selectedChart?.isPie ? 'appRecolorSlices' : 'appRecolorSeries'),
              colorOptions,
            )
          ) : (
            <RibbonButton
              large
              label={t('appChangeColors')}
              detail={t('appSelectEditableChart')}
              symbol="🎨"
              disabled
              onClick={() => {}}
            />
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appChangeChartType')}>
          {(
            [
              ['column', 'appChartColumn'],
              ['bar', 'appChartBar'],
              ['line', 'appChartLine'],
              ['area', 'appChartTypeArea'],
              ['pie', 'appChartPie'],
              ['doughnut', 'appChartDoughnut'],
            ] as const
          ).map(([type, labelKey]) => (
            <RibbonButton
              key={type}
              large
              label={t(labelKey)}
              detail={
                selectedChart?.currentType === type
                  ? t('appCurrentType')
                  : canConvert
                    ? t('appChangeChartType')
                    : t('appScatterKeepType')
              }
              symbol={type === 'line' || type === 'area' ? '🗠' : '📊'}
              active={selectedChart?.currentType === type}
              disabled={!canConvert}
              onClick={() => onCommand(`chart-type-${type}`)}
            />
          ))}
          {stackable
            ? largeMenu(t('appGroupingBtn'), '▤', t('appGroupingTitle'), groupingOptions)
            : null}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupData')}>
          <RibbonButton
            large
            label={t('appSelectDataBtn')}
            detail={t('appSeriesNamesRanges')}
            symbol="📊"
            disabled={!canEditChart}
            onClick={() => onCommand('chart-select-data')}
          />
          <RibbonButton
            large
            label={t('appSwitchRowColumn')}
            detail={t('appSwitchRowColumnDetail')}
            symbol="⇄"
            disabled={!canEditChart || (selectedChart?.categoryCount ?? 0) === 0}
            onClick={() => onCommand('chart-switch-row-col')}
          />
          <RibbonButton
            large
            label={t('appFormatPane')}
            detail={t('appFormatPaneDetail')}
            symbol="🎨"
            disabled={!canEditChart}
            onClick={() => onCommand('chart-format-pane')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupChartActions')}>
          <RibbonButton
            large
            label={t('appDeleteChart')}
            detail={selectedChart?.title ?? t('appRemoveFromSheet')}
            symbol="🗑"
            onClick={() => onCommand('chart-delete')}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Shape Format' || activeTab === 'Picture Format') {
    const paintSwatch = (color: string | undefined): string =>
      color === undefined || color === 'none' ? 'transparent' : color
    // No current paint: the glyph opens the palette instead of re-applying
    // a placeholder color.
    const hasPaint = (color: string | undefined): color is string =>
      color !== undefined && color !== 'none'
    const zOrder = (prefix: 'front' | 'back') => [
      {
        value: `visual:z:${prefix}`,
        label: t(prefix === 'front' ? 'appVisualBringToFront' : 'appVisualSendToBack'),
      },
      {
        value: `visual:z:${prefix === 'front' ? 'forward' : 'backward'}`,
        label: t(prefix === 'front' ? 'appVisualBringForward' : 'appVisualSendBackward'),
      },
    ]
    const alignOptions = [
      { value: 'visual:align:left', label: t('appVisualAlignLeft') },
      { value: 'visual:align:center', label: t('appVisualAlignCenter') },
      { value: 'visual:align:right', label: t('appVisualAlignRight') },
      { value: 'visual:align:top', label: t('appVisualAlignTop') },
      { value: 'visual:align:middle', label: t('appVisualAlignMiddle') },
      { value: 'visual:align:bottom', label: t('appVisualAlignBottom') },
      { value: 'visual:distribute:horizontal', label: t('appVisualDistributeH') },
      { value: 'visual:distribute:vertical', label: t('appVisualDistributeV') },
    ]
    const rotateOptions = [
      { value: 'visual:rotate:cw', label: t('appVisualRotateCw') },
      { value: 'visual:rotate:ccw', label: t('appVisualRotateCcw') },
      { value: 'visual:flip:v', label: t('appVisualFlipV') },
      { value: 'visual:flip:h', label: t('appVisualFlipH') },
    ]
    return (
      <div className="ribbon" data-ribbon-body="">
        {selectedShape && (
          <RibbonGroup label={t('appGroupShapeStyles')}>
            <ColorDropdown
              label="Shape fill"
              data-tip={t('appShapeFill')}
              split={hasPaint(selectedShape.fillColor)}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◧" />
                  <i style={{ background: paintSwatch(selectedShape.fillColor) }} />
                </span>
              }
              value={hasPaint(selectedShape.fillColor) ? selectedShape.fillColor : '#FFFFFF'}
              auto={t('dlgFcNoFill')}
              onPick={(hex) => onCommand(`shape-fill:${hex ?? 'none'}`)}
            />
            <ColorDropdown
              label="Shape outline"
              data-tip={t('appShapeOutline')}
              split={hasPaint(selectedShape.lineColor)}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◻" />
                  <i style={{ background: paintSwatch(selectedShape.lineColor) }} />
                </span>
              }
              value={hasPaint(selectedShape.lineColor) ? selectedShape.lineColor : '#FFFFFF'}
              auto={t('appShapeNoOutline')}
              onPick={(hex) => onCommand(`shape-outline:${hex ?? 'none'}`)}
            />
            <RibbonButton
              large
              label={t('appDeleteShape')}
              detail={t('appRemoveFromSheet')}
              symbol="🗑"
              onClick={() => onCommand('shape-delete')}
            />
          </RibbonGroup>
        )}
        <RibbonGroup label={t('appGroupArrange')}>
          {largeMenu(t('appVisualBringToFront'), '⬆', t('appVisualBringToFront'), zOrder('front'))}
          {largeMenu(t('appVisualSendToBack'), '⬇', t('appVisualSendToBack'), zOrder('back'))}
          {largeMenu(t('appVisualAlign'), '⫴', t('appVisualAlignTitle'), alignOptions)}
          {largeMenu(t('appVisualRotate'), '⟳', t('appVisualRotateTitle'), rotateOptions)}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSize')}>
          <RibbonButton
            large
            label={t('appVisualSizeProperties')}
            detail=""
            symbol="⤢"
            onClick={() => onCommand('visual:size')}
          />
          <RibbonButton
            large
            label={t('appVisualEditAltText')}
            detail=""
            symbol="Alt"
            onClick={() => onCommand('visual:alt-text')}
          />
        </RibbonGroup>
      </div>
    )
  }
  if (activeTab === 'Table Design') {
    const table = selectedTable
    const option = (key: TableOptionKey, label: string): React.JSX.Element => (
      <button
        className="check-item"
        data-tip={label}
        disabled={!table}
        onClick={() => onCommand(`table-option:${key}:${table?.[key] ? 0 : 1}`)}
      >
        <i className="check-box">{table?.[key] ? '\u2713' : ''}</i>
        {label}
      </button>
    )
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupTableProperties')}>
          <div className="row-stack table-name-field">
            <label className="table-name-label">{t('appTableNameLabel')}</label>
            <TableNameInput
              key={table?.key ?? ''}
              name={table?.name ?? ''}
              onCommit={(name) => onCommand(`table-rename:${name}`)}
            />
          </div>
          <RibbonButton
            large
            label={t('appResizeTable')}
            detail={table?.rangeA1 ?? ''}
            symbol="⤢"
            disabled={!table}
            onClick={() => onCommand('table-resize-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableTools')}>
          <RibbonButton
            large
            label={t('appRemoveDuplicates')}
            detail={t('appInSelection')}
            symbol="⧉"
            onClick={() => onCommand('remove-duplicates-open')}
          />
          <RibbonButton
            large
            label={t('appConvertToRange')}
            detail={table?.name ?? ''}
            symbol="▦"
            disabled={!table}
            onClick={() => onCommand('table-convert-range')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableStyleOptions')}>
          <div className="row-stack table-options">
            {option('headerRow', t('appTableOptHeaderRow'))}
            {option('totalsRow', t('appTableOptTotalRow'))}
            {option('bandedRows', t('appTableOptBandedRows'))}
          </div>
          <div className="row-stack table-options">
            {option('firstColumn', t('appTableOptFirstColumn'))}
            {option('lastColumn', t('appTableOptLastColumn'))}
            {option('bandedColumns', t('appTableOptBandedColumns'))}
          </div>
          <div className="row-stack table-options">
            {option('filterButton', t('appTableOptFilterButton'))}
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableStyles')}>
          <TableStyleGallery
            current={table?.style ?? null}
            onPick={(style) => onCommand(`table-style:${style}`)}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Insert') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupTables')}>
          <RibbonButton
            large
            label={t('appPivotTable')}
            detail={t('appFromSelection')}
            symbol="⊞"
            onClick={() => onCommand('pivot-open')}
          />
          <RibbonButton
            large
            label={t('appEditPivotTable')}
            detail={t('appChangeFields')}
            symbol="⊞"
            onClick={() => onCommand('pivot-edit')}
          />
          <RibbonButton
            large
            label={t('appTableBtn')}
            detail={t('appRealExcelTable')}
            symbol="▦"
            onClick={() => onCommand('format-as-table')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupIllustrations')}>
          <RibbonButton
            large
            label={t('appPictures')}
            detail={t('appPictureTypes')}
            symbol="🖼"
            onClick={() => onCommand('insert-picture')}
          />
          <div className="row-stack">
            <span className="styles-row" data-tip={t('appInsertShapeTitle')}>
              <ToolSymbol symbol="◇" />
              {t('appShapes')}
              <CaretIcon />
              <ShapeGallerySelect
                label="Shapes"
                onPick={(prst) => onCommand(`insert-shape:${prst}`)}
              />
            </span>
            <button
              className="styles-row as-button"
              data-tip={t('appIcons')}
              onClick={() => onCommand('insert-icons')}
            >
              <ToolSymbol symbol="✧" />
              {t('appIcons')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appScreenshot')}
              onClick={() => onCommand('insert-screenshot')}
            >
              <ToolSymbol symbol="⧉" />
              {t('appScreenshot')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCheckbox')}>
          <RibbonButton
            large
            label={t('appGroupCheckbox')}
            detail={t('appAtSelection')}
            symbol="☑"
            onClick={() => onCommand('insert-checkbox')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCharts')}>
          <RibbonButton
            large
            label={t('appRecommendedCharts')}
            detail={t('appFromSelection')}
            symbol="📊"
            onClick={() => onCommand('recommended-charts-open')}
          />
          <div className="chart-grid">
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartColumn') })}
              aria-label={t('appChartGridTitle', { type: t('appChartColumn') })}
              onClick={() => onCommand('insert-chart:column')}
            >
              <ToolSymbol symbol="▮▬" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartBar') })}
              aria-label={t('appChartGridTitle', { type: t('appChartBar') })}
              onClick={() => onCommand('insert-chart:bar')}
            >
              <ToolSymbol symbol="▤" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartLine') })}
              aria-label={t('appChartGridTitle', { type: t('appChartLine') })}
              onClick={() => onCommand('insert-chart:line')}
            >
              <ToolSymbol symbol="📈" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartTypeArea') })}
              aria-label={t('appChartGridTitle', { type: t('appChartTypeArea') })}
              onClick={() => onCommand('insert-chart:area')}
            >
              <ToolSymbol symbol="◪" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartPie') })}
              aria-label={t('appChartGridTitle', { type: t('appChartPie') })}
              onClick={() => onCommand('insert-chart:pie')}
            >
              <ToolSymbol symbol="◔" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartScatter') })}
              aria-label={t('appChartGridTitle', { type: t('appChartScatter') })}
              onClick={() => onCommand('insert-chart:scatter')}
            >
              <ToolSymbol symbol="∴" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartRadar') })}
              aria-label={t('appChartGridTitle', { type: t('appChartRadar') })}
              onClick={() => onCommand('insert-chart:radar')}
            >
              <ToolSymbol symbol="✳" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartDoughnut') })}
              aria-label={t('appChartGridTitle', { type: t('appChartDoughnut') })}
              onClick={() => onCommand('insert-chart:doughnut')}
            >
              <ToolSymbol symbol="◍" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartCombo') })}
              aria-label={t('appChartGridTitle', { type: t('appChartCombo') })}
              onClick={() => onCommand('insert-chart:combo')}
            >
              <ToolSymbol symbol="𝄜" />
            </button>
          </div>
          {largeMenu(
            t('appPivotChart'),
            '🗠',
            onIsSelectionInPivot() ? t('appPivotChartHintIn') : t('appPivotChartHintOut'),
            [
              { value: 'insert-pivot-chart:column', label: t('appChartColumn') },
              { value: 'insert-pivot-chart:bar', label: t('appChartBar') },
              { value: 'insert-pivot-chart:line', label: t('appChartLine') },
              { value: 'insert-pivot-chart:pie', label: t('appChartPie') },
              { value: 'insert-pivot-chart:doughnut', label: t('appChartDoughnut') },
              { value: 'insert-pivot-chart:radar', label: t('appChartRadar') },
            ],
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSparklines')}>
          {largeMenu(t('appGroupSparklines'), '〜', t('appSparklinesTitle'), [
            { value: 'sparkline:line', label: t('appSparkLine') },
            { value: 'sparkline:column', label: t('appSparkColumn') },
            { value: 'sparkline:stacked', label: t('appSparkWinLoss') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupFilters')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={onIsSelectionInPivot() ? t('appSlicerHintIn') : t('appSlicerHintOut')}
              onClick={() => onCommand('slicer-open')}
            >
              <ToolSymbol symbol="▥" />
              {t('appSlicer')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={onIsSelectionInPivot() ? t('appTimelineHintIn') : t('appTimelineHintOut')}
              onClick={() => onCommand('timeline-open')}
            >
              <ToolSymbol symbol="🕒" />
              {t('appTimeline')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupLinks')}>
          <RibbonButton
            large
            label={t('appLink')}
            detail={t('appAtSelection')}
            symbol="🔗"
            onClick={() => onCommand('link-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupComments')}>
          <RibbonButton
            large
            label={t('appNewComment')}
            detail={t('appAtSelection')}
            symbol="🗨"
            onClick={() => onCommand('note-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupText')}>
          <RibbonButton
            large
            label={t('appTextBox')}
            detail={t('appAtSelection')}
            symbol="A"
            onClick={() => onCommand('insert-textbox')}
          />
          <RibbonButton
            large
            label={t('appHeaderFooter')}
            detail={t('appPrintedPages')}
            symbol="🗎"
            onClick={() => onCommand('header-footer-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSymbols')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appEquation')}
              onClick={() => onCommand('insert-equation')}
            >
              <ToolSymbol symbol="π" />
              {t('appEquation')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appSymbol')}
              onClick={() => onCommand('insert-symbol')}
            >
              <ToolSymbol symbol="Ω" />
              {t('appSymbol')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Page Layout') {
    const fitOptions = (axis: 'width' | 'height') => {
      const current = String(
        (axis === 'width' ? pageLayout.fitToWidth : pageLayout.fitToHeight) ?? 0,
      )
      const options = [
        { value: '0', label: t('appFitAutomaticOption') },
        ...[1, 2, 3, 4, 5].map((pages) => ({
          value: String(pages),
          label: t(pages === 1 ? 'appFitPage1' : 'appFitPagesN', { count: pages }),
        })),
      ]
      return (
        <MenuSelect
          className="select-like compact"
          label={`Fit to ${axis}`}
          data-tip={axis === 'width' ? t('appWidth') : t('appHeight')}
          value={current}
          display={options.find((option) => option.value === current)?.label ?? current}
          options={options}
          onPick={(value) => onCommand(`page-layout:fit-${axis}:${value}`)}
        />
      )
    }
    const marginLabels = {
      normal: t('appMarginNormal'),
      wide: t('appMarginWide'),
      narrow: t('appMarginNarrow'),
    } as const
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupThemes')}>
          {largeMenu(
            t('appGroupThemes'),
            '🎨',
            t('appThemesTitle'),
            THEME_PRESETS.map((preset) => ({
              value: `page-layout:theme:${preset.id}`,
              label: preset.name,
            })),
          )}
          <div className="row-stack">
            <span className="styles-row" data-tip={t('appThemeColorsTitle')}>
              <ToolSymbol symbol="▤" />
              {t('appColors')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appColors')}
                options={COLOR_SCHEMES.map((scheme) => ({
                  value: `page-layout:theme-colors:${scheme.id}`,
                  label: scheme.name,
                }))}
                onPick={onCommand}
              />
            </span>
            <span className="styles-row" data-tip={t('appThemeFontsTitle')}>
              <ToolSymbol symbol="A" />
              {t('appFonts')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appFonts')}
                options={FONT_SCHEMES.map((scheme) => ({
                  value: `page-layout:theme-fonts:${scheme.id}`,
                  label: scheme.name,
                }))}
                onPick={onCommand}
              />
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupPageSetup')}>
          {largeMenu(
            t('appMargins'),
            '⿴',
            t('appMarginsTitle', {
              value: pageLayout.margins ? marginLabels[pageLayout.margins] : t('appAsSavedInFile'),
            }),
            [
              { value: 'page-layout:margins:normal', label: t('appMarginNormal') },
              { value: 'page-layout:margins:wide', label: t('appMarginWide') },
              { value: 'page-layout:margins:narrow', label: t('appMarginNarrow') },
            ],
          )}
          {largeMenu(
            t('appOrientationLabel'),
            '⤢',
            t('appOrientationTitle', {
              value: pageLayout.orientation
                ? t(pageLayout.orientation === 'portrait' ? 'appPortrait' : 'appLandscape')
                : t('appAsSavedInFile'),
            }),
            [
              { value: 'page-layout:orientation:portrait', label: t('appPortrait') },
              { value: 'page-layout:orientation:landscape', label: t('appLandscape') },
            ],
          )}
          {largeMenu(t('appSizeLabel'), '▭', t('appPaperSizeTitle'), [
            { value: 'page-layout:paper:1', label: 'Letter' },
            { value: 'page-layout:paper:5', label: 'Legal' },
            { value: 'page-layout:paper:3', label: 'Tabloid' },
            { value: 'page-layout:paper:7', label: 'Executive' },
            { value: 'page-layout:paper:8', label: 'A3' },
            { value: 'page-layout:paper:9', label: 'A4' },
            { value: 'page-layout:paper:11', label: 'A5' },
          ])}
          {largeMenu(
            t('appPrintArea'),
            '⬚',
            pageLayout.printArea
              ? t('appPrintAreaTitle', { area: pageLayout.printArea })
              : t('appPrintAreaFromSelection'),
            [
              { value: 'page-layout:print-area:set', label: t('appSetPrintArea') },
              { value: 'page-layout:print-area:clear', label: t('appClearPrintArea') },
            ],
          )}
          {largeMenu(t('appBreaks'), '┆', t('appBreaksTitle'), [
            { value: 'page-layout:breaks:insert', label: t('appInsertPageBreak') },
            { value: 'page-layout:breaks:remove', label: t('appRemovePageBreak') },
            { value: 'page-layout:breaks:reset', label: t('appResetAllPageBreaks') },
          ])}
          {largeMenu(
            t('appPrintTitlesLabel'),
            '▤',
            pageLayout.printTitles
              ? t('appPrintTitlesTitle', { rows: pageLayout.printTitles })
              : t('appPrintTitlesHint'),
            [
              { value: 'page-layout:print-titles:first-row', label: t('appRepeatRow1') },
              { value: 'page-layout:print-titles:selection', label: t('appRepeatSelectedRows') },
              { value: 'page-layout:print-titles:clear', label: t('appClearPrintTitles') },
            ],
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupScaleToFit')}>
          <div className="row-stack">
            <label className="styles-row" data-tip={t('appWidth')}>
              <ToolSymbol symbol="↔" />
              {t('appWidth')}:{fitOptions('width')}
            </label>
            <label className="styles-row" data-tip={t('appHeight')}>
              <ToolSymbol symbol="↕" />
              {t('appHeight')}:{fitOptions('height')}
            </label>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSheetOptions')}>
          <div className="check-column">
            <span className="check-head">{t('appGridlines')}</span>
            <button
              className="check-item"
              data-tip={t('appGridlines')}
              onClick={() => onCommand('toggle-gridlines')}
            >
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appViewCheck')}
            </button>
            <button
              className="check-item"
              data-tip={t('appPrintGridlinesTitle')}
              onClick={() =>
                onCommand(`page-layout:print-gridlines:${pageLayout.printGridlines ? '0' : '1'}`)
              }
            >
              <i className="check-box">{pageLayout.printGridlines ? '✓' : ''}</i>
              {t('appPrintCheck')}
            </button>
          </div>
          <div className="check-column">
            <span className="check-head">{t('appHeadings')}</span>
            <button
              className="check-item"
              data-tip={t('appHeadings')}
              onClick={() => onCommand('toggle-headings')}
            >
              <i className="check-box">{pageLayout.showHeadings ? '✓' : ''}</i>
              {t('appViewCheck')}
            </button>
            <button
              className="check-item"
              data-tip={t('appPrintHeadingsTitle')}
              onClick={() =>
                onCommand(`page-layout:print-headings:${pageLayout.printHeadings ? '0' : '1'}`)
              }
            >
              <i className="check-box">{pageLayout.printHeadings ? '✓' : ''}</i>
              {t('appPrintCheck')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Formulas') {
    const definedNames = onListNames()
    // Category buttons all open the same catalog dialog; the per-category
    // menus funnel into Insert Function.
    // Each button opens the catalog filtered to its own category; 'All'
    // for the ones the catalog has no counterpart for (Recently Used / More Functions).
    const functionCategory = (label: string, symbol: string, category: string) => (
      <RibbonButton
        large
        menu
        label={label}
        detail={category === 'All' ? t('appBrowseCatalog') : t('appBrowseCatalogFiltered')}
        symbol={symbol}
        onClick={() => onCommand(`insert-function-open:${category}`)}
      />
    )
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupFunctionLibrary')}>
          <RibbonButton
            large
            label={t('appInsertFunction')}
            detail={t('appBrowseCatalog')}
            symbol="ƒx"
            onClick={() => onCommand('insert-function-open')}
          />
          <div className="ribbon-tool large" data-tip={t('appAutoSumTitle')}>
            <span className="tool-icon-row">
              <ToolSymbol symbol="Σ" />
              <CaretIcon />
            </span>
            <span>
              <strong>{t('appAutoSum')}</strong>
            </span>
            <MenuSelect
              cover
              label="AutoSum"
              options={autoSumOptions(t)}
              onPick={(value) => onCommand(`autofn:${value}`)}
            />
          </div>
          {functionCategory(t('appFnCatRecentlyUsed'), '🕘', 'All')}
          {functionCategory(t('appFnCatFinancial'), '$', 'Financial')}
          {functionCategory(t('appFnCatLogical'), '?', 'Logical')}
          {functionCategory(t('appFnCatText'), 'A', 'Text')}
          {functionCategory(t('appFnCatDateTime'), '🕐', 'Date & Time')}
          {functionCategory(t('appFnCatLookup'), '🔍', 'Lookup')}
          {functionCategory(t('appFnCatMathTrig'), 'θ', 'Math')}
          {functionCategory(t('appFnCatMore'), '⋯', 'All')}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupDefinedNames')}>
          <RibbonButton
            large
            label={t('appNameManager')}
            detail={t('appNameManagerDetail')}
            symbol="🏷"
            onClick={() => onCommand('name-manager-open')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appDefineName')}
              onClick={() => onCommand('name-manager-open')}
            >
              <ToolSymbol symbol="🏷" />
              {t('appDefineName')}
              <CaretIcon />
            </button>
            <span className="styles-row" data-tip={t('appUseInFormulaTitle')}>
              <ToolSymbol symbol="ƒ" />
              {t('appUseInFormula')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appUseInFormula')}
                options={
                  definedNames.length > 0
                    ? definedNames.map((name) => ({ value: `use-in-formula:${name}`, label: name }))
                    : [{ value: 'name-manager-open', label: t('appNoNamesYet') }]
                }
                onPick={onCommand}
              />
            </span>
            <span className="styles-row" data-tip={t('appCreateFromSelectionTitle')}>
              <ToolSymbol symbol="⊞" />
              {t('appCreateFromSelection')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appCreateFromSelection')}
                options={[
                  { value: 'create-names:top', label: t('appCreateNamesTopRow') },
                  { value: 'create-names:left', label: t('appCreateNamesLeftCol') },
                ]}
                onPick={onCommand}
              />
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupFormulaAuditing')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appTracePrecedentsTitle')}
              onClick={() => onCommand('trace-precedents')}
            >
              <ToolSymbol symbol="⇢" />
              {t('appTracePrecedents')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appTraceDependentsTitle')}
              onClick={() => onCommand('trace-dependents')}
            >
              <ToolSymbol symbol="⇠" />
              {t('appTraceDependents')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appRemoveArrowsTitle')}
              onClick={() => onCommand('remove-arrows')}
            >
              <ToolSymbol symbol="⌫" />
              {t('appRemoveArrows')}
            </button>
          </div>
          <RibbonButton
            large
            label={t('appShowFormulas')}
            detail={t('appShowFormulasDetail')}
            symbol="ƒ"
            onClick={() => onCommand('toggle-show-formulas')}
          />
          <RibbonButton
            large
            label={t('appErrorChecking')}
            detail={t('appErrorCheckingDetail')}
            symbol="⚠"
            onClick={() => onCommand('error-checking')}
          />
          <RibbonButton
            large
            label={t('appWatchWindow')}
            detail={t('appWatchWindowDetail')}
            symbol="👓"
            onClick={() => onCommand('watch-window')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCalculation')}>
          {largeMenu(t('appCalculationOptions'), '🧮', t('appCalculationOptionsTitle'), [
            { value: 'calc-mode:auto', label: t('appCalcAuto') + (calcManual ? '' : ' ✓') },
            { value: 'calc-mode:manual', label: t('appCalcManual') + (calcManual ? ' ✓' : '') },
          ])}
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appCalculateNowTitle')}
              onClick={() => onCommand('calculate-now')}
            >
              <ToolSymbol symbol="⟳" />
              {t('appCalculateNow')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appCalculateSheetTitle')}
              onClick={() => onCommand('calculate-sheet')}
            >
              <ToolSymbol symbol="▦" />
              {t('appCalculateSheet')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Data') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appPivotTable')}>
          <RibbonButton
            large
            label={t('appPivotTable')}
            detail={t('appOoxmlPivot')}
            symbol="⊞"
            onClick={() => onCommand('pivot-open')}
          />
          <RibbonButton
            large
            label={t('appRefresh')}
            detail={onIsSelectionInPivot() ? t('appRefreshHintIn') : t('appRefreshHintOut')}
            symbol="⟳"
            onClick={() => {
              const err = onRefreshPivot()
              if (err) onCommand(`error:${err}`)
            }}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupGetData')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appFromTextCsvTitle')}
              onClick={() => onCommand('import-csv')}
            >
              <ToolSymbol symbol="🗎" />
              {t('appFromTextCsv')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appMergeWorkbooksTip')}
              onClick={() => onCommand('merge-workbooks')}
            >
              <ToolSymbol symbol="⧉" />
              {t('appMergeWorkbooks')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appRefreshAllTitle')}
              onClick={() => onCommand('refresh-all')}
            >
              <ToolSymbol symbol="⟳" />
              {t('appRefreshAll')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSortFilter')}>
          {largeMenu(t('appSort'), '⇅', t('appSortSelectionTitle'), [
            { value: 'sort:asc', label: t('appSortAToZ') },
            { value: 'sort:desc', label: t('appSortZToA') },
            { value: 'sort-custom-open', label: t('appCustomSort') },
          ])}
          <RibbonButton
            large
            label={t('appFilter')}
            detail={platformShortcuts(t('appFilterToggleDetail'))}
            symbol="▽"
            onClick={() => onCommand('filter-toggle')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appClearFilterTitle')}
              onClick={() => onCommand('filter-clear')}
            >
              <ToolSymbol symbol="⊘" />
              {t('appClear')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appReapplyTitle')}
              onClick={() => onCommand('filter-reapply')}
            >
              <ToolSymbol symbol="↻" />
              {t('appReapply')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appAdvancedFilterTitle')}
              onClick={() => onCommand('filter-advanced')}
            >
              <ToolSymbol symbol="▽" />
              {t('appAdvanced')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupDataTools')}>
          {largeMenu(t('appTextToColumns'), '⇶', t('appTextToColumnsTitle'), [
            { value: 'text-to-columns:2', label: t('appSplitByComma') },
            { value: 'text-to-columns:4', label: t('appSplitBySemicolon') },
            { value: 'text-to-columns:8', label: t('appSplitBySpace') },
            { value: 'text-to-columns:1', label: t('appSplitByTab') },
          ])}
          <RibbonButton
            large
            label={t('appFlashFill')}
            detail={t('appFlashFillDetail')}
            symbol="⚡"
            onClick={() => onCommand('flash-fill')}
          />
          <RibbonButton
            large
            label={t('appRemoveDuplicates')}
            detail={t('appInSelection')}
            symbol="⧉"
            onClick={() => onCommand('remove-duplicates-open')}
          />
          <RibbonButton
            large
            menu
            label={t('appDataValidation')}
            detail={t('appRulesPanel')}
            symbol="✓"
            onClick={() => onCommand('dv-open')}
          />
          <RibbonButton
            large
            label={t('appConsolidate')}
            detail={t('appCombineRanges')}
            symbol="⊕"
            onClick={() => onCommand('consolidate-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupForecast')}>
          {largeMenu(t('appWhatIfAnalysis'), '❔', t('appWhatIfTitle'), [
            { value: 'goal-seek-open', label: t('appGoalSeek') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupOutline')}>
          {largeMenu(t('appOutlineGroup'), '⊟', t('appOutlineGroupTitle'), [
            { value: 'outline-group:rows', label: t('appGroupRows') },
            { value: 'outline-group:cols', label: t('appGroupCols') },
            { value: 'outline-hide-detail:rows', label: t('appHideDetailRows') },
            { value: 'outline-hide-detail:cols', label: t('appHideDetailCols') },
            {
              value: 'outline-summary-below',
              label: `${outlineSummary.below ? '✓ ' : ''}${t('appOutlineSummaryBelow')}`,
            },
            {
              value: 'outline-summary-right',
              label: `${outlineSummary.right ? '✓ ' : ''}${t('appOutlineSummaryRight')}`,
            },
          ])}
          {largeMenu(t('appOutlineUngroup'), '⊞', t('appOutlineUngroupTitle'), [
            { value: 'outline-ungroup:rows', label: t('appUngroupRows') },
            { value: 'outline-ungroup:cols', label: t('appUngroupCols') },
            { value: 'outline-show-detail:rows', label: t('appShowDetailRows') },
            { value: 'outline-show-detail:cols', label: t('appShowDetailCols') },
            { value: 'outline-clear', label: t('appClearOutline') },
          ])}
          <RibbonButton
            large
            label={t('appSubtotal')}
            detail={t('appSubtotalDetail')}
            symbol="∑"
            onClick={() => onCommand('subtotal-open')}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'View') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupWorkbookViews')}>
          <RibbonButton
            large
            label={t('appNormalView')}
            detail={t('appCurrentViewTitle')}
            symbol="▦"
            active={pageLayout.pageBreakPreview !== true}
            onClick={() => {
              if (pageLayout.pageBreakPreview === true) onCommand('toggle-page-break-preview')
            }}
          />
          <RibbonButton
            large
            label={t('appPageBreakPreview')}
            detail={t('appPageBreakPreviewTitle')}
            symbol="┆"
            active={pageLayout.pageBreakPreview === true}
            onClick={() => {
              if (pageLayout.pageBreakPreview !== true) onCommand('toggle-page-break-preview')
            }}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupShow')}>
          <div className="check-column">
            <button
              className="check-item"
              data-tip={t('appGridlines')}
              onClick={() => onCommand('toggle-gridlines')}
            >
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appGridlines')}
            </button>
            <button
              className="check-item"
              data-tip={t('appFormulaBar')}
              onClick={() => onCommand('toggle-formula-bar')}
            >
              <i className="check-box">{formulaBarVisible ? '✓' : ''}</i>
              {t('appFormulaBar')}
            </button>
            <button
              className="check-item"
              data-tip={t('appHeadings')}
              onClick={() => onCommand('toggle-headings')}
            >
              <i className="check-box">{pageLayout.showHeadings ? '✓' : ''}</i>
              {t('appHeadings')}
            </button>
            <button
              className="check-item"
              data-tip={t('appCrossHighlight')}
              onClick={() => onCommand('toggle-cross-highlight')}
            >
              <i className="check-box">{crossHighlightVisible ? '✓' : ''}</i>
              {t('appCrossHighlight')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appZoomLabel')}>
          {largeMenu(t('appZoomLabel'), '🔍', t('appZoomSheetTitle'), [
            { value: 'zoom:50', label: '50%' },
            { value: 'zoom:75', label: '75%' },
            { value: 'zoom:100', label: '100%' },
            { value: 'zoom:125', label: '125%' },
            { value: 'zoom:150', label: '150%' },
            { value: 'zoom:200', label: '200%' },
          ])}
          <RibbonButton
            large
            label="100%"
            detail={t('appResetZoom')}
            symbol="⊙"
            onClick={() => onCommand('zoom-reset')}
          />
          <RibbonButton
            large
            label={t('appZoomToSelection')}
            detail={t('appZoomToSelectionDetail')}
            symbol="⌖"
            onClick={() => onCommand('zoom-to-selection')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupWindow')}>
          {largeMenu(t('appFreezePanes'), '❄', t('appFreezeTitle'), [
            { value: 'freeze-here', label: t('appFreezeAtSelection') },
            { value: 'freeze-top-row', label: t('appFreezeTopRow') },
            { value: 'freeze-first-col', label: t('appFreezeFirstCol') },
            { value: 'unfreeze', label: t('appUnfreeze') },
          ])}
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Review') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupProofing')}>
          <RibbonButton
            large
            label={t('appWorkbookStatsLabel')}
            detail={t('appSheetsCellsFormulas')}
            symbol="🧮"
            onClick={() => onCommand('workbook-statistics')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupLanguage')}>
          <div className="ribbon-tool large" data-tip={t('appTranslateTitle')}>
            <span className="tool-icon-row">
              <ToolSymbol symbol="文" />
              <CaretIcon />
            </span>
            <span>
              <strong>{t('appTranslate')}</strong>
            </span>
            <MenuSelect
              cover
              label={t('appTranslate')}
              options={TRANSLATE_LANGUAGES.map((language) => ({
                value: language,
                label: language,
              }))}
              onPick={(language) => onAiRun(t('appTranslatePrompt', { language }))}
            />
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupComments')}>
          <RibbonButton
            large
            label={t('appNewComment')}
            detail={t('appNewCommentDetail')}
            symbol="🗨"
            onClick={() => onCommand('comment-new')}
          />
          <RibbonButton
            large
            label={t('appDeleteLabel')}
            detail={t('appCommentAtSelection')}
            symbol="🗑"
            onClick={() => onCommand('comment-delete')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appNotePrevTitle')}
              onClick={() => onCommand('comment-prev')}
            >
              <ToolSymbol symbol="←" />
              {t('appPrevious')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appNoteNextTitle')}
              onClick={() => onCommand('comment-next')}
            >
              <ToolSymbol symbol="→" />
              {t('appNext')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appShowCommentsPaneTitle')}
              onClick={() => onCommand('comments-pane-toggle')}
            >
              <ToolSymbol symbol="🗨" />
              {t('appShowComments')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupNotes')}>
          {largeMenu(t('appGroupNotes'), '🗒', t('appNotesDetail'), [
            { value: 'note-open', label: t('appNewNote') },
            { value: 'note-show-toggle', label: t('appShowHideNote') },
            { value: 'note-show-all', label: t('appShowAllNotes') },
            { value: 'notes-convert', label: t('appConvertToNotes') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupProtection')}>
          <RibbonButton
            large
            label={t(sheetProtected ? 'appUnprotectSheet' : 'appProtectSheet')}
            detail={t(sheetProtected === null ? 'appOpenFileFirst' : 'appProtectSheetDetail')}
            symbol={sheetProtected ? '🔓' : '🔒'}
            onClick={() => onCommand('sheet-protect')}
          />
          <RibbonButton
            large
            label={t(workbookProtected ? 'appUnprotectWorkbook' : 'appProtectWorkbook')}
            detail={t(workbookProtected === null ? 'appOpenFileFirst' : 'appProtectWorkbookTitle')}
            symbol={workbookProtected ? '🔓' : '🔐'}
            onClick={() => onCommand('workbook-protect')}
          />
          <RibbonButton
            large
            label={t('appAllowEditRanges')}
            detail={t('appAllowEditRangesTitle')}
            symbol="⬚"
            onClick={() => onCommand('allow-edit-ranges-open')}
          />
        </RibbonGroup>
      </div>
    )
  }

  const fontSizes = [9, 10, 11, 12, 14, 16, 18, 22, 26]
  const echoFamily = selectionFormat?.fontFamily ?? 'Aptos'
  const echoSize = selectionFormat?.fontSize ?? 11
  const fontGroups = fontFamilyGroups(systemFontFamilies, echoFamily)
  const familyOptions = [
    ...fontGroups.common.map((family) => ({ value: family, label: family })),
    ...fontGroups.system.map((family, index) => ({
      value: family,
      label: family,
      sep: index === 0,
    })),
  ]
  const sizeOptions = fontSizes.includes(echoSize)
    ? fontSizes
    : [...fontSizes, echoSize].sort((a, b) => a - b)
  return (
    <div className="ribbon" data-ribbon-body="">
      <RibbonGroup label={t('appGroupAiAssistant')}>
        <button
          className={`ribbon-tool as-button large ai-entry ${aiOpen ? 'active' : ''}`}
          data-tip={t('aiOpenAssistant')}
          onClick={onAiToggle}
        >
          <span className="tool-icon-row">
            <GensparkMark size={26} />
          </span>
          <span>
            <strong>Genspark AI</strong>
          </span>
        </button>
        <button
          className="ribbon-tool as-button large ai-entry"
          disabled={!sheetHasContent}
          data-tip={t('aiCheckBtn')}
          onClick={() => onAiRun(t('aiCheckPrompt'))}
        >
          <span className="tool-icon-row">
            <span className="ai-feature-icon" aria-hidden="true">
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M11 3.25C15.2802 3.25 18.75 6.71979 18.75 11C18.75 15.2802 15.2802 18.75 11 18.75C6.71979 18.75 3.25 15.2802 3.25 11C3.25 6.71979 6.71979 3.25 11 3.25Z" />
                <path
                  d="M7.5 10.8235L9.64097 12.9645C9.93755 13.2611 10.4177 13.2634 10.7171 12.9697L14.7647 9"
                  strokeLinecap="round"
                />
                <path d="M20 20.5L16.5 17" strokeLinecap="round" />
              </svg>
            </span>
          </span>
          <span>
            <strong>{t('aiCheckBtn')}</strong>
          </span>
        </button>
        <button
          className="ribbon-tool as-button large ai-entry"
          disabled={!sheetHasContent}
          data-tip={t('aiAnalyzeBtn')}
          onClick={() => onAiRun(t('aiAnalyzePrompt'))}
        >
          <span className="tool-icon-row">
            <span className="ai-feature-icon" aria-hidden="true">
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M3.88589 14.2073H8.48682" strokeLinecap="round" />
                <path d="M3.88589 19.0112H8.48682" strokeLinecap="round" />
                <path d="M3.88589 9.40369H11.692" strokeLinecap="round" />
                <path d="M3.88589 4.59998H19.1645" strokeLinecap="round" />
                <path d="M15.1995 10.5445C15.3784 10.0908 16.0206 10.0908 16.1996 10.5445L16.706 11.8286C17.0338 12.6598 17.6918 13.3178 18.523 13.6456L19.8071 14.1521C20.2608 14.331 20.2608 14.9732 19.8071 15.1522L18.523 15.6586C17.6918 15.9864 17.0338 16.6444 16.706 17.4756L16.1996 18.7597C16.0206 19.2134 15.3784 19.2134 15.1995 18.7597L14.693 17.4756C14.3652 16.6444 13.7072 15.9864 12.876 15.6586L11.592 15.1522C11.1382 14.9732 11.1382 14.331 11.592 14.1521L12.876 13.6456C13.7072 13.3178 14.3652 12.6598 14.693 11.8286L15.1995 10.5445Z" />
              </svg>
            </span>
          </span>
          <span>
            <strong>{t('aiAnalyzeBtn')}</strong>
          </span>
        </button>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupClipboard')}>
        <button
          className="ribbon-tool as-button large"
          data-tip={t('appPasteTitle')}
          onClick={() => onCommand('paste')}
        >
          <span className="tool-icon-row">
            <ToolSymbol symbol="📋" />
          </span>
          <span>
            <strong>{t('appPaste')}</strong>
          </span>
        </button>
        {largeMenu(t('appPasteSpecial'), '📑', t('appPasteSpecialTitle'), [
          { value: 'paste-special:value', label: t('appPasteValuesOnly') },
          { value: 'paste-special:formula', label: t('appPasteFormulasOnly') },
          { value: 'paste-special:format', label: t('appPasteFormattingOnly') },
          { value: 'paste-special:col-width', label: t('appPasteColWidths') },
          { value: 'paste-special:besides-border', label: t('appPasteExceptBorders') },
          { value: PASTE_SPECIAL_OPEN_SHELL_COMMAND, label: t('appPasteSpecialDialog') },
        ])}
        <div className="tool-stack">
          <button
            data-tip={t('appCutTitle')}
            aria-label={t('appCutTitle')}
            onClick={() => onCommand('cut')}
          >
            <ToolSymbol symbol="✂" />
          </button>
          <button
            data-tip={t('appCopyTitle')}
            aria-label={t('appCopyTitle')}
            onClick={() => onCommand('copy')}
          >
            <ToolSymbol symbol="⧉" />
          </button>
          <button
            data-tip={t('appFormatPainterTip')}
            aria-label={t('appFormatPainter')}
            aria-pressed={formatPainterActive}
            className={formatPainterActive ? 'is-active' : ''}
            onClick={(event) =>
              onCommand(event.detail >= 2 ? 'format-painter:dblclick' : 'format-painter')
            }
          >
            <ToolSymbol symbol="🖌" />
          </button>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupFont')}>
        <div className="ribbon-rows">
          <div className="inline-tools">
            <EditableMenuSelect
              className="select-like font-name"
              label="Font family"
              data-tip={t('appFontFamilyTip')}
              value={echoFamily}
              options={familyOptions}
              onOpen={loadSystemFonts}
              onPick={(value) => onCommand(`font-family:${value}`)}
              commit={(text) => onCommand(`font-family:${text}`)}
            />
            <EditableMenuSelect
              className="select-like font-size"
              label="Font size"
              data-tip={t('appFontSizeTip')}
              value={String(echoSize)}
              options={sizeOptions.map((size) => ({ value: String(size), label: String(size) }))}
              onPick={(value) => onCommand(`font-size:${value}`)}
              commit={(text) => {
                const size = Number(text.replace(',', '.'))
                // Excel's font-size bounds
                if (Number.isFinite(size) && size >= 1 && size <= 409)
                  onCommand(`font-size:${size}`)
              }}
            />
            <button
              data-tip={t('appIncreaseFontSize')}
              aria-label={t('appIncreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, 1)}`)}
            >
              <ToolSymbol symbol="A↑" />
            </button>
            <button
              data-tip={t('appDecreaseFontSize')}
              aria-label={t('appDecreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, -1)}`)}
            >
              <ToolSymbol symbol="A↓" />
            </button>
          </div>
          <div className="inline-tools">
            <button
              data-tip={t('appBold')}
              className={selectionFormat?.bold ? 'is-active' : ''}
              onClick={() => onCommand('bold')}
            >
              <b>B</b>
            </button>
            <button
              data-tip={t('appItalic')}
              className={selectionFormat?.italic ? 'is-active' : ''}
              onClick={() => onCommand('italic')}
            >
              <em>I</em>
            </button>
            <button
              data-tip={t('appUnderline')}
              className={selectionFormat?.underline ? 'is-active' : ''}
              onClick={() => onCommand('underline')}
            >
              <u>U</u>
            </button>
            <button
              data-tip={t('appDoubleUnderline')}
              onClick={() => onCommand('underline:double')}
            >
              <u style={{ textDecorationStyle: 'double' }}>D</u>
            </button>
            <button
              data-tip={t('appStrikethrough')}
              className={selectionFormat?.strike ? 'is-active' : ''}
              onClick={() => onCommand('strike')}
            >
              <s>S</s>
            </button>
            <button
              data-tip={t('appSuperscript')}
              aria-label={t('appSuperscript')}
              className={selectionFormat?.vertAlign === 'superscript' ? 'is-active' : ''}
              onClick={() => onCommand('superscript')}
            >
              <ToolSymbol symbol="x²" />
            </button>
            <button
              data-tip={t('appSubscript')}
              aria-label={t('appSubscript')}
              className={selectionFormat?.vertAlign === 'subscript' ? 'is-active' : ''}
              onClick={() => onCommand('subscript')}
            >
              <ToolSymbol symbol="x₂" />
            </button>
            <ColorDropdown
              label="Font color"
              data-tip={t('appFontColor')}
              split
              display={
                <span className="swatch-letter">
                  A<i style={{ background: fontColor }} />
                </span>
              }
              value={fontColor}
              auto={t('appAutomaticColor')}
              onPick={(hex) => {
                setFontColor(hex ?? '#000000')
                onCommand(hex ? `font-color:${hex}` : 'font-color:auto')
              }}
            />
            <ColorDropdown
              label="Fill color"
              data-tip={t('appFillColor')}
              split
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◧" />
                  <i style={{ background: fillColor }} />
                </span>
              }
              value={fillColor}
              auto={t('dlgFcNoFill')}
              onPick={(hex) => {
                if (hex) setFillColor(hex)
                onCommand(hex ? `fill:${hex}` : 'fill:none')
              }}
            />
            <MenuSelect
              className="select-like compact"
              label="Borders"
              data-tip={t('appBorders')}
              display={
                <>
                  <ToolSymbol symbol="⊡" /> {t('appBorders')}
                </>
              }
              options={[
                { value: 'all', label: t('appBorderAll'), icon: <BorderAllIcon /> },
                { value: 'outer', label: t('appBorderOuter'), icon: <BorderOuterIcon /> },
                {
                  value: 'thick-outer',
                  label: t('appBorderThickOuter'),
                  icon: <BorderThickOuterIcon />,
                },
                { value: 'top', label: t('appBorderTop'), icon: <BorderTopIcon /> },
                { value: 'bottom', label: t('appBorderBottom'), icon: <BorderBottomIcon /> },
                { value: 'left', label: t('appBorderLeft'), icon: <BorderLeftIcon /> },
                { value: 'right', label: t('appBorderRight'), icon: <BorderRightIcon /> },
                { value: 'none', label: t('appBorderNone'), icon: <BorderNoneIcon /> },
              ]}
              onPick={(value) => onCommand(`border:${value}:${borderColor}`)}
            />
            <ColorDropdown
              label="Border color"
              data-tip={t('appBorderColor')}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="⊡" />
                  <i style={{ background: borderColor }} />
                </span>
              }
              value={borderColor}
              onPick={(hex) => {
                if (hex) setBorderColor(hex)
              }}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupAlignment')}>
        <div className="ribbon-rows">
          <div className="inline-tools alignment-tools">
            <button
              data-tip={t('appTopAlign')}
              aria-label={t('appTopAlign')}
              onClick={() => onCommand('valign:top')}
            >
              <ToolSymbol symbol="⤒" />
            </button>
            <button
              data-tip={t('appMiddleAlign')}
              aria-label={t('appMiddleAlign')}
              onClick={() => onCommand('valign:middle')}
            >
              <ToolSymbol symbol="↕" />
            </button>
            <button
              data-tip={t('appBottomAlign')}
              aria-label={t('appBottomAlign')}
              onClick={() => onCommand('valign:bottom')}
            >
              <ToolSymbol symbol="⤓" />
            </button>
            <button
              data-tip={t('dlgFcWrapText')}
              aria-label={t('dlgFcWrapText')}
              className={selectionFormat?.wrap ? 'is-active' : ''}
              onClick={() => onCommand('wrap')}
            >
              <ToolSymbol symbol="↩" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Orientation"
              data-tip={t('dlgFcOrientation')}
              display={<ToolSymbol symbol="⤴" />}
              options={[
                { value: '45', label: t('appAngleCcw') },
                { value: '-45', label: t('appAngleCw') },
                { value: 'vertical', label: t('appVerticalText') },
                { value: '90', label: t('appRotateUp') },
                { value: '-90', label: t('appRotateDown') },
                { value: '0', label: t('appClearRotation') },
              ]}
              onPick={(value) => onCommand(`rotate:${value}`)}
            />
          </div>
          <div className="inline-tools alignment-tools">
            <button
              data-tip={t('appAlignLeft')}
              aria-label={t('appAlignLeft')}
              onClick={() => onCommand('align:left')}
            >
              <ToolSymbol symbol="≡" />
            </button>
            <button
              data-tip={t('appAlignCenter')}
              aria-label={t('appAlignCenter')}
              onClick={() => onCommand('align:center')}
            >
              <ToolSymbol symbol="≣" />
            </button>
            <button
              data-tip={t('appAlignRight')}
              aria-label={t('appAlignRight')}
              onClick={() => onCommand('align:right')}
            >
              <ToolSymbol symbol="☰" />
            </button>
            <button
              data-tip={t('appDecreaseIndent')}
              aria-label={t('appDecreaseIndent')}
              disabled={!selectionFormat?.indent}
              onClick={() => onCommand('indent-step:-1')}
            >
              <ToolSymbol symbol="⇤" />
            </button>
            <button
              data-tip={t('appIncreaseIndent')}
              aria-label={t('appIncreaseIndent')}
              disabled={(selectionFormat?.indent ?? 0) >= 250}
              onClick={() => onCommand('indent-step:1')}
            >
              <ToolSymbol symbol="⇥" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Merge cells"
              data-tip={t('appMergeCells')}
              display={
                <>
                  <ToolSymbol symbol="⇔" /> {t('appMerge')}
                </>
              }
              options={[
                { value: 'center', label: t('appMergeCenter') },
                { value: 'across', label: t('appMergeAcross') },
                { value: 'cells', label: t('appMergeCells') },
                { value: 'unmerge', label: t('appUnmergeCells') },
              ]}
              onPick={(value) => onCommand(`merge:${value}`)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupNumber')}>
        <div className="ribbon-rows">
          <NumberFormatSelect pattern={selectionFormat?.numberFormat ?? ''} onCommand={onCommand} />
          <div className="inline-tools">
            <button
              data-tip={t('appCurrency')}
              aria-label={t('appCurrency')}
              onClick={() => onCommand(`format:${accountingStyle(RIBBON_CURRENCY_SYMBOL)}`)}
            >
              $
            </button>
            <button
              data-tip={t('appPercentTitle')}
              aria-label={t('appPercentTitle')}
              onClick={() => onCommand(`format:${PERCENT_STYLE}`)}
            >
              %
            </button>
            <button
              data-tip={t('appThousands')}
              aria-label={t('appThousands')}
              onClick={() => onCommand(`format:${COMMA_STYLE}`)}
            >
              ,
            </button>
            <button
              data-tip={t('appIncreaseDecimal')}
              aria-label={t('appIncreaseDecimal')}
              onClick={() => onCommand('decimal-inc')}
            >
              .0+
            </button>
            <button
              data-tip={t('appDecreaseDecimal')}
              aria-label={t('appDecreaseDecimal')}
              onClick={() => onCommand('decimal-dec')}
            >
              .0−
            </button>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupStyles')}>
        <div className="styles-stack">
          <button
            className="styles-row as-button"
            data-tip={t('appConditionalFormatting')}
            onClick={() => onCommand('cf-open')}
          >
            <ToolSymbol symbol="▤" />
            {t('appConditionalFormatting')}
            <CaretIcon />
          </button>
          <div className="styles-row as-button" data-tip={t('appFormatAsTableTitle')}>
            <ToolSymbol symbol="▦" />
            {t('appFormatAsTable')}
            <CaretIcon />
            <MenuSelect
              cover
              label="Format as Table"
              options={[
                { value: 'TableStyleLight1', label: t('appTableStyleLight1') },
                { value: 'TableStyleLight9', label: t('appTableStyleLight9') },
                { value: 'TableStyleMedium2', label: t('appTableStyleMedium2') },
                { value: 'TableStyleMedium4', label: t('appTableStyleMedium4') },
                { value: 'TableStyleMedium7', label: t('appTableStyleMedium7') },
                { value: 'TableStyleDark2', label: t('appTableStyleDark2') },
              ]}
              onPick={(value) => onCommand(`format-as-table:${value}`)}
            />
          </div>
          <div className="styles-row as-button" data-tip={t('appCellStylesTitle')}>
            <ToolSymbol symbol="🎨" />
            {t('appCellStyles')}
            <CaretIcon />
            <MenuSelect
              cover
              label="Cell Styles"
              options={[
                { value: 'good', label: t('appStyleGood') },
                { value: 'bad', label: t('appStyleBad') },
                { value: 'neutral', label: t('appStyleNeutral') },
                { value: 'input', label: t('appStyleInput') },
                { value: 'output', label: t('appStyleOutput') },
                { value: 'calculation', label: t('appStyleCalculation') },
                { value: 'warning-text', label: t('appStyleWarningText') },
                { value: 'title', label: t('appTitle') },
                { value: 'heading-1', label: t('appStyleHeading1') },
                { value: 'heading-2', label: t('appStyleHeading2') },
                { value: 'total', label: t('appStyleTotal') },
                { value: 'accent1-20', label: t('appStyleAccent120') },
                { value: 'accent1-40', label: t('appStyleAccent140') },
                { value: 'accent1', label: t('appStyleAccent1') },
              ]}
              onPick={(value) => onCommand(`cell-style:${value}`)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupCells')}>
        <div className="ribbon-rows">
          <button
            className="styles-row as-button"
            data-tip={t('appFormatCells')}
            data-tip-kbd={platformShortcuts('⌘1')}
            onClick={() => onCommand('format-cells')}
          >
            <ToolSymbol symbol="🎨" />
            {t('appFormatCells')} {platformShortcuts('⌘1')}
            <CaretIcon />
          </button>
          <div className="inline-tools cell-tools">
            <button
              data-tip={t('appInsertRow')}
              aria-label={t('appInsertRow')}
              onClick={() => onCommand('insert-row-here')}
            >
              <ToolSymbol symbol="⤒" />
            </button>
            <button
              data-tip={t('appDeleteRow')}
              aria-label={t('appDeleteRow')}
              onClick={() => onCommand('delete-row-here')}
            >
              <ToolSymbol symbol="⤓" />
            </button>
            <button
              data-tip={t('appInsertCol')}
              aria-label={t('appInsertCol')}
              onClick={() => onCommand('insert-col-here')}
            >
              <ToolSymbol symbol="⇤" />
            </button>
            <button
              data-tip={t('appDeleteCol')}
              aria-label={t('appDeleteCol')}
              onClick={() => onCommand('delete-col-here')}
            >
              <ToolSymbol symbol="⇥" />
            </button>
            <button
              data-tip={t('appInsertCells')}
              data-tip-kbd={platformShortcuts('⇧⌘=')}
              aria-label={t('appInsertCells')}
              onClick={() => onCommand('insert-cells')}
            >
              <ToolSymbol symbol="⊞" />
            </button>
            <button
              data-tip={t('appDeleteCells')}
              data-tip-kbd={platformShortcuts('⌘-')}
              aria-label={t('appDeleteCells')}
              onClick={() => onCommand('delete-cells')}
            >
              <ToolSymbol symbol="⊟" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Format"
              data-tip={t('appFormatMenu')}
              display={
                <>
                  <ToolSymbol symbol="⇳" /> {t('appFormatMenu')}
                </>
              }
              options={[
                { value: 'row-height-open', label: `${t('appRowHeight')}…` },
                { value: 'autofit-row-height', label: t('appAutoFitRowHeight') },
                { value: 'col-width-open', label: `${t('appColWidth')}…` },
                { value: 'autofit-col-width', label: t('appAutoFitColWidth') },
              ]}
              onPick={(value) => onCommand(value)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupEditing')}>
        <div className="ribbon-rows">
          <div className="inline-tools">
            <MenuSelect
              className="select-like compact"
              label="AutoSum"
              data-tip={t('appAutoSumTitle')}
              display={
                <>
                  <ToolSymbol symbol="Σ" /> {t('appAutoSum')}
                </>
              }
              options={autoSumOptions(t)}
              onPick={(value) => onCommand(`autofn:${value}`)}
            />
            <MenuSelect
              className="select-like compact"
              label="Sort & Filter"
              data-tip={t('appGroupSortFilter')}
              display={
                <>
                  <ToolSymbol symbol="⇅" /> {t('appGroupSortFilter')}
                </>
              }
              options={[
                { value: 'sort:asc', label: t('appSortAToZ') },
                { value: 'sort:desc', label: t('appSortZToA') },
                { value: 'sort-custom-open', label: t('appCustomSort') },
                { value: 'filter-toggle', label: t('appFilter') },
                { value: 'filter-clear', label: t('appClearFilterTitle') },
                { value: 'filter-reapply', label: t('appReapplyTitle') },
              ]}
              onPick={(value) => onCommand(value)}
            />
            <MenuSelect
              className="select-like compact"
              label="Fill"
              data-tip={t('appFillMenu')}
              display={
                <>
                  <ToolSymbol symbol="↓" /> {t('appFillMenu')}
                </>
              }
              options={[
                { value: 'fill-down', label: t('appFillDown') },
                { value: 'fill-right', label: t('appFillRight') },
                { value: 'fill-up', label: t('appFillUp') },
                { value: 'fill-left', label: t('appFillLeft') },
                { value: 'fill-series-open', label: t('appFillSeries') },
                { value: 'flash-fill', label: t('appFlashFill') },
              ]}
              onPick={(value) => onCommand(value)}
            />
            <MenuSelect
              className="select-like compact"
              label="Clear"
              data-tip={t('appClear')}
              display={
                <>
                  <ToolSymbol symbol="⌫" /> {t('appClear')}
                </>
              }
              options={[
                { value: 'all', label: t('appClearAll') },
                { value: 'formats', label: t('appClearFormats') },
                { value: 'contents', label: t('appClearContents') },
                { value: 'comments', label: t('appClearComments') },
                { value: 'hyperlinks', label: t('appClearHyperlinks') },
              ]}
              onPick={(value) => onCommand(`clear-${value}`)}
            />
          </div>
          <div className="inline-tools">
            <button
              className="labeled"
              data-tip={t('appFindTitle')}
              onClick={() => onCommand('find')}
            >
              <ToolSymbol symbol="🔍" /> {t('appFind')}
            </button>
            <button
              className="labeled"
              data-tip={t('appReplace')}
              onClick={() => onCommand('replace')}
            >
              <ToolSymbol symbol="⇄" /> {t('appReplace')}
            </button>
            <button
              className="labeled"
              data-tip={`${t('appGoTo')} ${platformShortcuts('⌘G')}`}
              onClick={() => onCommand('goto-open')}
            >
              <ToolSymbol symbol="⌖" /> {t('appGoTo')}
            </button>
          </div>
        </div>
      </RibbonGroup>
    </div>
  )
}

function autoSumOptions(t: (key: StringKey) => string): { value: string; label: string }[] {
  return [
    { value: 'SUM', label: t('appFnSum') },
    { value: 'AVERAGE', label: t('appFnAverage') },
    { value: 'COUNT', label: t('appFnCountNumbers') },
    { value: 'MAX', label: t('appFnMax') },
    { value: 'MIN', label: t('appFnMin') },
  ]
}

const STAT_LABEL: Record<StatusBarFunc, StringKey> = {
  AVERAGE: 'appStatAverage',
  COUNTA: 'appStatCount',
  COUNT: 'appStatNumericalCount',
  MIN: 'appStatMin',
  MAX: 'appStatMax',
  SUM: 'appStatSum',
}

/// Excel's status bar right-click menu: tick the aggregates the footer shows.
function StatusBarStatsMenu({
  at,
  enabled,
  onToggle,
  onClose,
}: {
  readonly at: { x: number; y: number }
  readonly enabled: readonly StatusBarFunc[]
  readonly onToggle: (func: StatusBarFunc) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  useEscapeClose(true, onClose)
  useEffect(() => {
    const onPress = (event: MouseEvent): void => {
      if (!ref.current?.contains(event.target as Node)) onClose()
    }
    window.addEventListener('mousedown', onPress)
    return () => window.removeEventListener('mousedown', onPress)
  }, [onClose])
  // Flip up when the cursor sits near the bottom: the status bar always does.
  const style: React.CSSProperties = {
    left: Math.min(at.x, window.innerWidth - 200),
    bottom: Math.max(8, window.innerHeight - at.y),
  }
  return (
    <div ref={ref} className="status-stats-menu" role="menu" style={style}>
      {STATUS_BAR_FUNCS.map((func) => (
        <button
          key={func}
          type="button"
          role="menuitemcheckbox"
          aria-checked={enabled.includes(func)}
          onClick={() => onToggle(func)}
        >
          <i className="check-box">{enabled.includes(func) ? '✓' : ''}</i>
          {t(STAT_LABEL[func])}
        </button>
      ))}
    </div>
  )
}

/// Escape-to-close for the ribbon dropdowns; outside-press / blur / shell
/// chrome-press dismissal lives in the shared useDismissablePopover.
function useEscapeClose(open: boolean, close: () => void): void {
  const closeRef = useRef(close)
  closeRef.current = close
  useEffect(() => {
    if (!open) return
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') closeRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open])
}

/// Custom ribbon dropdown replacing the native <select>: macOS pops the native
/// menu over the control (covering ribbon content), while this panel is
/// anchored below its trigger — same pattern as the slides ribbon's .rb-drop.
function MenuSelect({
  label,
  'data-tip': tip,
  className,
  cover = false,
  display,
  value = '',
  options,
  onPick,
}: {
  /// aria-label (mirrors the old select's aria-label)
  readonly label: string
  readonly 'data-tip'?: string
  /// trigger classes for the combobox look (e.g. 'select-like compact')
  readonly className?: string
  /// invisible trigger stretched over the host tool (old .cover-select)
  readonly cover?: boolean
  /// content shown in the trigger (combobox variants)
  readonly display?: React.ReactNode
  /// currently applied value, highlighted in the open panel ('' = none)
  readonly value?: string
  readonly options: readonly { value: string; label: string; icon?: React.ReactNode }[]
  readonly onPick: (value: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  return (
    <div ref={wrapRef} className={`menu-select${cover ? ' menu-select-cover' : ''}`}>
      <button
        type="button"
        className={cover ? 'cover-select' : className}
        data-tip={tip}
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        {!cover && (
          <>
            <span className="menu-select-value">{display}</span>
            <CaretIcon />
          </>
        )}
      </button>
      {open && (
        <div className="menu-select-drop" role="listbox" aria-label={label}>
          {options.map((option) => (
            <button
              type="button"
              key={option.value}
              role="option"
              aria-selected={value !== '' && option.value === value}
              className={value !== '' && option.value === value ? 'on' : ''}
              onClick={() => {
                setOpen(false)
                onPick(option.value)
              }}
            >
              {option.icon && <span className="menu-option-icon">{option.icon}</span>}
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

/// Cross-app shared shape gallery (slides parity), minus Lines — the grid
/// renderer draws shapes as filled paths and cannot show stroke-only connectors.
const SHEET_SHAPE_GROUPS = SHAPE_GALLERY_GROUPS.filter(
  (g) => g.groupKey !== 'ribbonShapeGroupLines',
)

/// Shapes dropdown: grouped outline-icon grid (same look/content as docs and
/// slides); the trigger is an invisible cover like MenuSelect's `cover` mode.
function ShapeGallerySelect({
  label,
  onPick,
}: {
  readonly label: string
  readonly onPick: (prst: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  return (
    <div ref={wrapRef} className="menu-select menu-select-cover">
      <button
        type="button"
        className="cover-select"
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      />
      {open && (
        <div className="menu-select-drop rb-shape-gallery" role="listbox" aria-label={label}>
          {SHEET_SHAPE_GROUPS.map((group) => (
            <div key={group.groupKey}>
              <div className="rb-drop-title">{t(group.groupKey as StringKey)}</div>
              <div className="rb-shape-grid">
                {group.shapes.map((s) => (
                  <button
                    type="button"
                    key={s.prst}
                    className="rb-shape-cell"
                    data-tip={t(s.labelKey as StringKey)}
                    aria-label={t(s.labelKey as StringKey)}
                    onClick={() => {
                      setOpen(false)
                      onPick(s.prst)
                    }}
                  >
                    <ShapePreview prst={s.prst} size={18} />
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

/// MenuSelect variant whose trigger is a free-text input (Excel's font
/// name/size boxes): Enter or clicking away commits the typed value, Esc
/// reverts, the caret opens the preset list.
function EditableMenuSelect({
  label,
  'data-tip': tip,
  className,
  value,
  options,
  onOpen,
  onPick,
  commit,
}: {
  readonly label: string
  readonly 'data-tip'?: string
  readonly className?: string
  readonly value: string
  readonly options: readonly { value: string; label: string; sep?: boolean }[]
  /// fired on the click that opens the list (lazy option loading)
  readonly onOpen?: () => void
  readonly onPick: (value: string) => void
  /// apply typed text (caller validates; invalid input is dropped silently)
  readonly commit: (text: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const [draft, setDraftState] = useState<string | null>(null)
  // mirrors draft synchronously: Enter/Esc call blur(), whose handler runs
  // before the setState above lands
  const draftRef = useRef<string | null>(null)
  const setDraft = (next: string | null): void => {
    draftRef.current = next
    setDraftState(next)
  }
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  const commitDraft = (): void => {
    const text = draftRef.current?.trim()
    if (text && text !== value) commit(text)
    setDraft(null)
  }
  return (
    <div ref={wrapRef} className="menu-select">
      <span className={`${className ?? ''} menu-select-edit`} data-tip={tip}>
        <input
          value={draft ?? value}
          aria-label={label}
          onFocus={(event) => event.currentTarget.select()}
          onChange={(event) => setDraft(event.currentTarget.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.currentTarget.blur() // blur handler commits
            } else if (event.key === 'Escape') {
              setDraft(null)
              event.currentTarget.blur()
            }
          }}
          onBlur={(event) => {
            // focus moved into the dropdown → let the option click win
            if (wrapRef.current?.contains(event.relatedTarget as Node)) return
            commitDraft()
          }}
        />
        <button
          type="button"
          className="menu-select-caret"
          aria-label={label}
          aria-haspopup="listbox"
          aria-expanded={open}
          onClick={() => {
            if (!open) onOpen?.()
            setOpen(!open)
          }}
        >
          <CaretIcon />
        </button>
      </span>
      {open && (
        <div className="menu-select-drop" role="listbox" aria-label={label}>
          {options.map((option) => (
            <button
              type="button"
              key={option.value}
              role="option"
              aria-selected={option.value === value}
              className={`${option.value === value ? 'on' : ''}${option.sep ? ' sep-above' : ''}`}
              onClick={() => {
                setOpen(false)
                setDraft(null)
                onPick(option.value)
              }}
            >
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

function NumberFormatSelect({
  pattern,
  onCommand,
}: {
  readonly pattern: string
  readonly onCommand: (command: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const current = categoryOptionForPattern(pattern)
  const localized = (label: string): string => {
    const labelKey = NUMBER_FORMAT_LABEL[label]
    return labelKey ? t(labelKey) : label
  }
  return (
    <MenuSelect
      className="select-like"
      label="Number format"
      data-tip={pattern || t('dlgFcNumGeneral')}
      value={current}
      display={localized(current)}
      options={numberFormatCategories().map((category) => ({
        value: category.label,
        label: localized(category.label),
      }))}
      onPick={(value) => {
        const category = numberFormatCategories().find((candidate) => candidate.label === value)
        if (category) onCommand(`format:${category.pattern}`)
      }}
    />
  )
}

function RibbonGroup({
  label,
  children,
}: {
  readonly label: string
  readonly children: React.ReactNode
}): React.JSX.Element {
  // No visible group captions — the label stays for assistive tech.
  return (
    <section className="ribbon-group" aria-label={label}>
      <div className="ribbon-group-content">{children}</div>
    </section>
  )
}

function RibbonButton({
  label,
  detail,
  symbol,
  accent = false,
  menu = false,
  compact = false,
  large = false,
  active = false,
  disabled = false,
  onClick,
}: {
  readonly label: string
  readonly detail: string
  readonly symbol: string
  readonly accent?: boolean
  readonly menu?: boolean
  /** Icon-only in narrow windows (e.g. Cut/Copy). */
  readonly compact?: boolean
  /** Icon over label, for the tall Insert-tab buttons. */
  readonly large?: boolean
  /** Pressed-state echo (e.g. highlighting the current chart type). */
  readonly active?: boolean
  readonly disabled?: boolean
  readonly onClick: () => void
}): React.JSX.Element {
  return (
    <button
      className={`ribbon-tool as-button ${accent ? 'accent' : ''} ${compact ? 'compact-icon' : ''} ${large ? 'large' : ''} ${active ? 'active' : ''}`}
      onClick={onClick}
      disabled={disabled}
      data-tip={label}
      data-tip-detail={detail}
    >
      {large ? (
        // The dropdown caret sits beside the icon (top row, inside the
        // hover plate), not after the label
        <span className="tool-icon-row">
          <ToolSymbol symbol={symbol} />
          {menu && <CaretIcon />}
        </span>
      ) : (
        <ToolSymbol symbol={symbol} />
      )}
      <span>
        <strong>
          {label}
          {!large && menu && <CaretIcon />}
        </strong>
        <small>{detail}</small>
      </span>
    </button>
  )
}

function fillSeriesDefaultType(numberFormat: string | undefined): FillSeriesType {
  return numberFormat && numberFormatLabel(numberFormat) === 'Date' ? 'date' : 'linear'
}
