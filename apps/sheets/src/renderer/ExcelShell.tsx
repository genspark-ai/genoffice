import { useEffect, useRef, useState } from 'react'

import {
  CaretIcon,
  GensparkMark,
  RIBBON_GLYPH_ICONS,
  RedoIcon,
  SaveIcon,
  UndoIcon,
} from './ribbon-icons'

import { FormatCellsDialog } from './FormatCellsDialog'
import { GoToDialog } from './GoToDialog'
import { useI18n, type StringKey } from './i18n/locale'
import { NameManagerDialog, type DefinedNameAction, type DefinedNameRow } from './NameManagerDialog'
import { categoryOptionForPattern, NUMBER_FORMAT_CATEGORIES } from './number-format'
import { type SelectionFormat } from './selection-format'

import type { ChartSeriesVisualState } from '../domain/chart-visual'
import type { ChangePlan } from '../domain/workbook.types'
import type { AttachmentMeta } from '../shared/desktop-api'
import { AiChatPanel, type AiChatMessage } from './ai/AiChatPanel'
import {
  PivotDialog,
  type PivotEditSeed,
  type PivotField,
  type OoXmlPivotConfig,
} from './PivotDialog'
import { InsertFunctionDialog } from './InsertFunctionDialog'
import { SubtotalDialog, type SubtotalConfig } from './SubtotalDialog'
import { ConsolidateDialog } from './ConsolidateDialog'
import type { ConsolidateConfig } from './consolidate'
import { HeaderFooterDialog, type HeaderFooterResult } from './HeaderFooterDialog'
import type { HeaderFooterParts } from './edit-journal'

// No File tab: file commands live in the macOS
// application menu (File → Open/Save/Save As) and the toolbar icons.
const ribbonTabs = ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'] as const

/// 'Chart Design' is contextual: it exists only while a chart is selected,
/// and appears without stealing the active tab.
type RibbonTab = (typeof ribbonTabs)[number] | 'Chart Design'

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

/// Excel's grow/shrink font walks its size ladder, not ±1.
const FONT_SIZE_LADDER = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 26, 28, 36, 48, 72]

function stepFontSize(current: number, direction: 1 | -1): number {
  if (direction === 1) {
    return FONT_SIZE_LADDER.find((size) => size > current) ?? FONT_SIZE_LADDER.at(-1) ?? current
  }
  return (
    [...FONT_SIZE_LADDER].reverse().find((size) => size < current) ?? FONT_SIZE_LADDER[0] ?? current
  )
}

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
  readonly prompt: string
  readonly preview: ChangePlan | null
  readonly selectionFormat: SelectionFormat | null
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
  /** Send the composer text, or the given instruction when provided */
  readonly onSend: (instruction?: string) => void
  readonly onStop: () => void
  readonly onNewChat: () => void
  readonly onUndo: () => void
  readonly onCommand: (command: string) => void
  /// Left side of the status bar (ready / streaming / AI progress messages).
  readonly statusMessage: string
  /// Zoom of the active sheet in percent, echoed by the status-bar slider.
  readonly zoomPercent: number
  /// True when the edit journal has unsaved changes (enables the QAT Save).
  readonly canSave: boolean
  readonly onSave: () => void
  /// QAT redo (workbook history, same path as the app menu's ⇧⌘Z); undo
  /// shares the AI panel's onUndo above.
  readonly onRedo: () => void
  /// AutoSave toggle in the tab row (docs/slides parity).
  readonly autoSave: boolean
  readonly onAutoSaveChange: (on: boolean) => void
  /// Non-null while a floating chart is selected in the grid.
  readonly selectedChart: SelectedChartRibbon | null
  /// Column choices of the active selection, read when the Sort dialog opens.
  readonly onGetSortColumns: () => { label: string; colIndex: number }[]
  /// Effective protection of the active sheet (null = unknown / demo).
  readonly onGetSheetProtection: () => boolean | null
  /// Name Manager data + actions (actions return an error message or null).
  readonly onGetDefinedNames: () => {
    names: DefinedNameRow[]
    sheets: { id: string; name: string }[]
  }
  readonly onDefinedNameAction: (action: DefinedNameAction) => string | null
  /// Field choices for the Pivot dialog, read from the selection's header row.
  readonly onGetPivotFields: () => PivotField[]
  readonly onGetSourceRange: () => string
  readonly onCreatePivot: (config: OoXmlPivotConfig) => string | null
  /// A3 editing of an existing pivot: when it returns null, App has already shown
  /// the reason and the dialog is not opened.
  readonly onGetPivotEditSeed: () => PivotEditSeed | null
  readonly onEditPivot: (config: OoXmlPivotConfig) => string | null
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
  readonly onApplyFormula: (formula: string) => string | null
  readonly onCreateSubtotal: (config: SubtotalConfig) => string | null
  readonly onCreateConsolidate: (config: ConsolidateConfig) => string | null
  /// Prefill for the Consolidate reference input (current multi-cell selection).
  readonly onGetConsolidateDefault: () => string
  /// Header & Footer dialog OK; returns an error message, or null on success.
  readonly onApplyHeaderFooter: (result: HeaderFooterResult) => string | null
  /// Session page-layout settings of the active sheet, echoed by the Page
  /// Layout tab's controls (untouched fields show the app default).
  readonly pageLayout: PageLayoutEcho
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
  readonly printArea?: string | null | undefined
  readonly printTitles?: string | null | undefined
  readonly header?: HeaderFooterParts | null | undefined
  readonly footer?: HeaderFooterParts | null | undefined
}

export function ExcelShell({
  prompt,
  preview,
  selectionFormat,
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
  onGetDefinedNames,
  onDefinedNameAction,
  onGetPivotFields,
  onGetSourceRange,
  onCreatePivot,
  onGetPivotEditSeed,
  onEditPivot,
  onRefreshPivot,
  onIsSelectionInPivot,
  onGetActiveCell,
  onGetAnchorValue,
  activeCellA1,
  onGoToReference,
  onListDefinedNames,
  onApplyFormula,
  onCreateSubtotal,
  onCreateConsolidate,
  onGetConsolidateDefault,
  onApplyHeaderFooter,
  onPromptChange,
  onSend,
  onStop,
  onNewChat,
  onUndo,
  onCommand,
  statusMessage,
  zoomPercent,
  canSave,
  onSave,
  onRedo,
  autoSave,
  onAutoSaveChange,
  selectedChart,
  pageLayout,
}: ExcelShellProps): React.JSX.Element {
  const { t } = useI18n()
  const [activeTab, setActiveTab] = useState<RibbonTab>('Home')
  const [isCopilotOpen, setIsCopilotOpen] = useState(true)
  const [showFormatCells, setShowFormatCells] = useState(false)
  const [axisSizeTarget, setAxisSizeTarget] = useState<'row' | 'col' | null>(null)
  const [showLinkDialog, setShowLinkDialog] = useState(false)
  const [showSortDialog, setShowSortDialog] = useState(false)
  const [showDedupeDialog, setShowDedupeDialog] = useState(false)
  const [showNameManager, setShowNameManager] = useState(false)
  const [showPivotDialog, setShowPivotDialog] = useState(false)
  const [pivotEditSeed, setPivotEditSeed] = useState<PivotEditSeed | null>(null)
  /** null = closed; string = open on that catalog category ('All' for the plain button) */
  const [insertFunctionCat, setInsertFunctionCat] = useState<string | null>(null)
  const [showSubtotalDialog, setShowSubtotalDialog] = useState(false)
  const [showConsolidateDialog, setShowConsolidateDialog] = useState(false)
  const [showGoTo, setShowGoTo] = useState(false)
  const [showHeaderFooter, setShowHeaderFooter] = useState(false)
  /// Non-null while the Chart Design → Add Chart Element text prompt is open.
  const [chartTextTarget, setChartTextTarget] = useState<ChartTextTarget | null>(null)
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent): void => {
      if ((event.metaKey || event.ctrlKey) && event.key === '1') {
        event.preventDefault()
        setShowFormatCells(true)
      }
      // Excel's Go To shortcut (⌘G / Ctrl+G).
      if ((event.metaKey || event.ctrlKey) && event.key === 'g') {
        event.preventDefault()
        setShowGoTo(true)
      }
      // Excel's Show Formulas shortcut (⌘` / Ctrl+`).
      if ((event.metaKey || event.ctrlKey) && event.key === '`') {
        event.preventDefault()
        onCommand('toggle-show-formulas')
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [onCommand])
  // Deselecting while on the contextual tab lands back on Home.
  useEffect(() => {
    if (!selectedChart && activeTab === 'Chart Design') setActiveTab('Home')
  }, [selectedChart, activeTab])
  const visibleTabs: readonly RibbonTab[] = selectedChart
    ? [...ribbonTabs, 'Chart Design']
    : ribbonTabs

  return (
    <main className={`app-shell ${isCopilotOpen ? '' : 'copilot-collapsed'}`}>
      <header className="excel-header">
        <nav
          className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}`}
          aria-label="Workbook commands"
        >
          <button
            type="button"
            className="qa-btn"
            title={t('appSaveTitle')}
            aria-label={t('appSaveTitle')}
            disabled={!canSave}
            onClick={onSave}
          >
            <SaveIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            title={t('appUndo')}
            aria-label={t('appUndo')}
            onClick={onUndo}
          >
            <UndoIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            title={t('appRedo')}
            aria-label={t('appRedo')}
            onClick={onRedo}
          >
            <RedoIcon />
          </button>
          <label className={`autosave-toggle ${autoSave ? 'on' : ''}`} title={t('appAutoSaveTip')}>
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
              className={`${tab === activeTab ? 'active' : ''} ${tab === 'Chart Design' ? 'contextual' : ''}`}
              key={tab}
              onClick={() => setActiveTab(tab)}
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
          sheetHasContent={sheetHasContent}
          sheetProtected={onGetSheetProtection()}
          pageLayout={pageLayout}
          selectedChart={selectedChart}
          onRefreshPivot={onRefreshPivot}
          onIsSelectionInPivot={onIsSelectionInPivot}
          onCommand={(command) => {
            if (command === 'format-cells') setShowFormatCells(true)
            else if (command === 'row-height-open') setAxisSizeTarget('row')
            else if (command === 'col-width-open') setAxisSizeTarget('col')
            else if (command === 'link-open') setShowLinkDialog(true)
            else if (command === 'sort-custom-open') setShowSortDialog(true)
            else if (command === 'remove-duplicates-open') setShowDedupeDialog(true)
            else if (command === 'name-manager-open') setShowNameManager(true)
            else if (command === 'pivot-open') setShowPivotDialog(true)
            else if (command === 'pivot-edit') setPivotEditSeed(onGetPivotEditSeed())
            else if (command === 'insert-function-open') setInsertFunctionCat('All')
            else if (command.startsWith('insert-function-open:'))
              setInsertFunctionCat(command.slice('insert-function-open:'.length))
            else if (command === 'subtotal-open') setShowSubtotalDialog(true)
            else if (command === 'consolidate-open') setShowConsolidateDialog(true)
            else if (command === 'goto-open') setShowGoTo(true)
            else if (command === 'header-footer-open') setShowHeaderFooter(true)
            else if (command === 'ai-open-panel') setIsCopilotOpen(true)
            else if (command === 'ai-toggle-panel') setIsCopilotOpen((v) => !v)
            else if (command === 'chart-element-title') setChartTextTarget('title')
            else if (command === 'chart-element-axis-cat') setChartTextTarget('axis-category')
            else if (command === 'chart-element-axis-val') setChartTextTarget('axis-value')
            else onCommand(command)
          }}
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
          onStop={onStop}
          onNewChat={onNewChat}
          onUndo={onUndo}
          onExpand={() => setIsCopilotOpen(true)}
          onCollapse={() => setIsCopilotOpen(false)}
        />
        <div className="sheet-main">
          {/* Excel's formula-bar row, Name Box only for now (fx bar TBD). */}
          <div className="name-box-bar">
            <NameBox activeCellA1={activeCellA1} onGoTo={onGoToReference} />
            <button
              className="name-box-goto"
              title={t('appGoToButtonTitle')}
              aria-label="Go To"
              onClick={() => setShowGoTo(true)}
            >
              ▾
            </button>
          </div>
          <section className="workbook-area">
            <div id="univer-container" className="spreadsheet" />
          </section>
        </div>
      </div>

      {/* Full-width status bar, unified with docs/slides (Univer's own zoom slider is disabled). */}
      <footer className="status-bar">
        <div className="status-left">
          <span className="status-msg">{statusMessage}</span>
        </div>
        <div className="status-right">
          <button className="zoom-btn" onClick={() => onCommand('zoom-out')}>
            −
          </button>
          <input
            className="zoom-slider"
            type="range"
            min={50}
            max={400}
            value={Math.min(400, Math.max(50, zoomPercent))}
            onChange={(event) => onCommand(`zoom:${event.target.value}`)}
          />
          <button className="zoom-btn" onClick={() => onCommand('zoom-in')}>
            +
          </button>
          <span className="zoom-value">{zoomPercent}%</span>
        </div>
      </footer>
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
      {showPivotDialog && (
        <PivotDialog
          fields={onGetPivotFields()}
          sourceRange={onGetSourceRange()}
          onCreate={onCreatePivot}
          onClose={() => setShowPivotDialog(false)}
        />
      )}
      {pivotEditSeed && (
        <PivotDialog
          mode="edit"
          fields={pivotEditSeed.fields}
          sourceRange={pivotEditSeed.sourceRange}
          initial={pivotEditSeed.initial}
          onCreate={onEditPivot}
          onClose={() => setPivotEditSeed(null)}
        />
      )}
      {insertFunctionCat !== null && (
        <InsertFunctionDialog
          targetLabel={onGetActiveCell()}
          onApply={onApplyFormula}
          initialCategory={insertFunctionCat}
          onClose={() => setInsertFunctionCat(null)}
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
      {showGoTo && (
        <GoToDialog
          names={onListDefinedNames()}
          onGo={onGoToReference}
          onClose={() => setShowGoTo(false)}
        />
      )}
      {showHeaderFooter && (
        <HeaderFooterDialog
          initialHeader={pageLayout.header ?? null}
          initialFooter={pageLayout.footer ?? null}
          onApply={onApplyHeaderFooter}
          onClose={() => setShowHeaderFooter(false)}
        />
      )}
    </main>
  )
}

/// Excel's Name Box: echoes the active cell while idle; focusing it starts a
/// draft, Enter jumps to the typed address or defined name (an invalid one
/// keeps the draft and flags the input), Esc or blur cancels back to the
/// echo. The echo prop updates via the SelectionChanged refresh in App.
function NameBox({
  activeCellA1,
  onGoTo,
}: {
  readonly activeCellA1: string
  readonly onGoTo: (ref: string) => string | null
}): React.JSX.Element {
  const { t } = useI18n()
  const [draft, setDraft] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  return (
    <input
      className={`name-box${error === null ? '' : ' invalid'}`}
      aria-label="Name Box"
      title={error ?? t('appNameBoxTitle')}
      placeholder="A1"
      spellCheck={false}
      value={draft ?? activeCellA1}
      onFocus={(event) => {
        setDraft(activeCellA1)
        event.target.select()
      }}
      onChange={(event) => {
        setDraft(event.target.value)
        setError(null)
      }}
      onBlur={() => {
        setDraft(null)
        setError(null)
      }}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          const failure = onGoTo(draft ?? activeCellA1)
          setError(failure)
          if (failure === null) {
            setDraft(null)
            event.currentTarget.blur()
          }
        } else if (event.key === 'Escape') {
          setDraft(null)
          setError(null)
          event.currentTarget.blur()
        }
      }}
    />
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
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        aria-label="Custom sort"
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
              <select
                className="select-like"
                value={level.colIndex}
                onChange={(event) => setLevel(index, { colIndex: Number(event.target.value) })}
              >
                {index > 0 && <option value={NO_SORT_LEVEL}>{t('appSortNone')}</option>}
                {columns.map((column) => (
                  <option key={column.colIndex} value={column.colIndex}>
                    {column.label}
                  </option>
                ))}
              </select>
              <select
                className="select-like compact"
                value={level.order}
                onChange={(event) => setLevel(index, { order: event.target.value as 'a' | 'd' })}
              >
                <option value="a">{t('appSortAsc')}</option>
                <option value="d">{t('appSortDesc')}</option>
              </select>
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
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        aria-label="Remove duplicates"
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
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
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
                if (event.key === 'Escape') onClose()
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
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
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
                if (event.key === 'Escape') onClose()
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
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        aria-label="Insert link"
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
                if (event.key === 'Escape') onClose()
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
  sheetHasContent,
  sheetProtected,
  pageLayout,
  selectedChart,
  onCommand,
  onAiRun,
  aiOpen,
  onAiToggle,
  onRefreshPivot,
  onIsSelectionInPivot,
}: {
  readonly activeTab: RibbonTab
  readonly selectionFormat: SelectionFormat | null
  readonly sheetHasContent: boolean
  readonly sheetProtected: boolean | null
  readonly pageLayout: PageLayoutEcho
  readonly selectedChart: SelectedChartRibbon | null
  readonly onCommand: (command: string) => void
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
  // Large menu button: a native select stretched invisibly over the tool,
  // each option carrying its full command string.
  const largeMenu = (
    label: string,
    symbol: string,
    title: string,
    options: readonly { value: string; label: string }[],
  ): React.JSX.Element => (
    <div className="ribbon-tool large" title={title}>
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
      <div className="ribbon">
        <RibbonGroup label={t('appGroupChartLayouts')}>
          {canEditChart ? (
            largeMenu(t('appAddChartElement'), '📊', t('appAddChartElementTitle'), elementOptions)
          ) : (
            <RibbonReserved large menu label={t('appAddChartElement')} symbol="📊" />
          )}
          {canEditChart ? (
            largeMenu(t('appQuickLayout'), '▦', t('appQuickLayoutTitle'), layoutOptions)
          ) : (
            <RibbonReserved large menu label={t('appQuickLayout')} symbol="▦" />
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

  if (activeTab === 'Insert') {
    return (
      <div className="ribbon">
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
        <RibbonGroup label={t('appGroupForms')}>
          <RibbonReserved large menu label={t('appGroupForms')} symbol="🗒" />
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
            <span className="styles-row" title={t('appInsertShapeTitle')}>
              <ToolSymbol symbol="◇" />
              {t('appShapes')}
              <CaretIcon />
              <MenuSelect
                cover
                label="Shapes"
                options={[
                  { value: 'rect', label: t('appShapeRect') },
                  { value: 'roundRect', label: t('appShapeRoundRect') },
                  { value: 'ellipse', label: t('appShapeEllipse') },
                  { value: 'triangle', label: t('appShapeTriangle') },
                  { value: 'diamond', label: t('appShapeDiamond') },
                  { value: 'rightArrow', label: t('appShapeRightArrow') },
                  { value: 'leftArrow', label: t('appShapeLeftArrow') },
                  { value: 'pentagon', label: t('appShapePentagon') },
                  { value: 'hexagon', label: t('appShapeHexagon') },
                ]}
                onPick={(value) => onCommand(`insert-shape:${value}`)}
              />
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="✧" />
              {t('appIcons')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="⬡" />
              {t('app3dModels')}
              <CaretIcon />
            </span>
          </div>
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="▤" />
              SmartArt
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="⧉" />
              {t('appScreenshot')}
              <CaretIcon />
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCheckbox')}>
          <RibbonReserved large label={t('appGroupCheckbox')} symbol="☑" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCharts')}>
          <RibbonReserved large label={t('appRecommendedCharts')} symbol="📊" />
          <div className="chart-grid">
            <button
              title={t('appChartGridTitle', { type: t('appChartColumn') })}
              onClick={() => onCommand('insert-chart:column')}
            >
              <ToolSymbol symbol="▮▬" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartBar') })}
              onClick={() => onCommand('insert-chart:bar')}
            >
              <ToolSymbol symbol="▤" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartLine') })}
              onClick={() => onCommand('insert-chart:line')}
            >
              <ToolSymbol symbol="📈" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartTypeArea') })}
              onClick={() => onCommand('insert-chart:area')}
            >
              <ToolSymbol symbol="◪" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartPie') })}
              onClick={() => onCommand('insert-chart:pie')}
            >
              <ToolSymbol symbol="◔" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartScatter') })}
              onClick={() => onCommand('insert-chart:scatter')}
            >
              <ToolSymbol symbol="∴" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartRadar') })}
              onClick={() => onCommand('insert-chart:radar')}
            >
              <ToolSymbol symbol="✳" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartDoughnut') })}
              onClick={() => onCommand('insert-chart:doughnut')}
            >
              <ToolSymbol symbol="◍" />
            </button>
            <button
              title={t('appChartGridTitle', { type: t('appChartCombo') })}
              onClick={() => onCommand('insert-chart:combo')}
            >
              <ToolSymbol symbol="𝄜" />
            </button>
          </div>
          <RibbonReserved large label={t('appMaps')} symbol="🌐" />
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
              title={onIsSelectionInPivot() ? t('appSlicerHintIn') : t('appSlicerHintOut')}
              onClick={() => onCommand('slicer-open')}
            >
              <ToolSymbol symbol="▥" />
              {t('appSlicer')}
            </button>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="🕒" />
              {t('appTimeline')}
            </span>
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
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="π" />
              {t('appEquation')}
              <CaretIcon />
            </span>
            <button className="styles-row as-button" onClick={() => onCommand('insert-symbol')}>
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
      <div className="ribbon">
        <RibbonGroup label={t('appGroupThemes')}>
          <RibbonReserved large menu label={t('appGroupThemes')} symbol="🎨" />
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="▤" />
              {t('appColors')}
              <CaretIcon />
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="A" />
              {t('appFonts')}
              <CaretIcon />
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
          <RibbonReserved large menu label={t('appBreaks')} symbol="┆" />
          <RibbonReserved large label={t('appBackground')} symbol="🖼" />
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
            <label className="styles-row">
              <ToolSymbol symbol="↔" />
              {t('appWidth')}:{fitOptions('width')}
            </label>
            <label className="styles-row">
              <ToolSymbol symbol="↕" />
              {t('appHeight')}:{fitOptions('height')}
            </label>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSheetOptions')}>
          <div className="check-column">
            <span className="check-head">{t('appGridlines')}</span>
            <button className="check-item" onClick={() => onCommand('toggle-gridlines')}>
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appViewCheck')}
            </button>
            <button
              className="check-item"
              title={t('appPrintGridlinesTitle')}
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
            <span className="check-item reserved-check" title={t('appNotAvailableYet')}>
              <i className="check-box">✓</i>
              {t('appViewCheck')}
            </span>
            <button
              className="check-item"
              title={t('appPrintHeadingsTitle')}
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
      <div className="ribbon">
        <RibbonGroup label={t('appGroupFunctionLibrary')}>
          <RibbonButton
            large
            label={t('appInsertFunction')}
            detail={t('appBrowseCatalog')}
            symbol="ƒx"
            onClick={() => onCommand('insert-function-open')}
          />
          <div className="ribbon-tool large" title={t('appAutoSumTitle')}>
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
              options={[
                { value: 'SUM', label: t('appFnSum') },
                { value: 'AVERAGE', label: t('appFnAverage') },
                { value: 'COUNT', label: t('appFnCountNumbers') },
                { value: 'MAX', label: t('appFnMax') },
                { value: 'MIN', label: t('appFnMin') },
              ]}
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
            <button className="styles-row as-button" onClick={() => onCommand('name-manager-open')}>
              <ToolSymbol symbol="🏷" />
              {t('appDefineName')}
              <CaretIcon />
            </button>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="ƒ" />
              {t('appUseInFormula')}
              <CaretIcon />
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="⊞" />
              {t('appCreateFromSelection')}
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupFormulaAuditing')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              title={t('appTracePrecedentsTitle')}
              onClick={() => onCommand('trace-precedents')}
            >
              <ToolSymbol symbol="⇢" />
              {t('appTracePrecedents')}
            </button>
            <button
              className="styles-row as-button"
              title={t('appTraceDependentsTitle')}
              onClick={() => onCommand('trace-dependents')}
            >
              <ToolSymbol symbol="⇠" />
              {t('appTraceDependents')}
            </button>
            <button
              className="styles-row as-button"
              title={t('appRemoveArrowsTitle')}
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
          <RibbonReserved large menu label={t('appErrorChecking')} symbol="⚠" />
          <RibbonReserved large label={t('appWatchWindow')} symbol="👓" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCalculation')}>
          <RibbonReserved large menu label={t('appCalculationOptions')} symbol="🧮" />
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appRecalcLiveTitle')}>
              <ToolSymbol symbol="⟳" />
              {t('appCalculateNow')}
            </span>
            <span className="styles-row reserved" title={t('appRecalcLiveTitle')}>
              <ToolSymbol symbol="▦" />
              {t('appCalculateSheet')}
            </span>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Data') {
    return (
      <div className="ribbon">
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
          <RibbonReserved large menu label={t('appGroupGetData')} symbol="🛢" />
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="🗎" />
              {t('appFromTextCsv')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="⟳" />
              {t('appRefreshAll')}
              <CaretIcon />
            </span>
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
            detail={t('appFilterToggleDetail')}
            symbol="▽"
            onClick={() => onCommand('filter-toggle')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              title={t('appClearFilterTitle')}
              onClick={() => onCommand('filter-clear')}
            >
              <ToolSymbol symbol="⊘" />
              {t('appClear')}
            </button>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="↻" />
              {t('appReapply')}
            </span>
            <button
              className="styles-row as-button"
              title={t('appAdvancedFilterTitle')}
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
          <RibbonReserved large menu label={t('appWhatIfAnalysis')} symbol="❔" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupOutline')}>
          {largeMenu(t('appOutlineGroup'), '⊟', t('appOutlineGroupTitle'), [
            { value: 'outline-group:rows', label: t('appGroupRows') },
            { value: 'outline-group:cols', label: t('appGroupCols') },
            { value: 'outline-hide-detail:rows', label: t('appHideDetailRows') },
            { value: 'outline-hide-detail:cols', label: t('appHideDetailCols') },
          ])}
          {largeMenu(t('appOutlineUngroup'), '⊞', t('appOutlineUngroupTitle'), [
            { value: 'outline-ungroup:rows', label: t('appUngroupRows') },
            { value: 'outline-ungroup:cols', label: t('appUngroupCols') },
            { value: 'outline-show-detail:rows', label: t('appShowDetailRows') },
            { value: 'outline-show-detail:cols', label: t('appShowDetailCols') },
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
      <div className="ribbon">
        <RibbonGroup label={t('appGroupWorkbookViews')}>
          <div className="ribbon-tool large is-current" title={t('appCurrentViewTitle')}>
            <span className="tool-icon-row">
              <ToolSymbol symbol="▦" />
            </span>
            <span>
              <strong>{t('appNormalView')}</strong>
            </span>
          </div>
          <RibbonReserved large label={t('appTabPageLayout')} symbol="🗎" />
          <RibbonReserved large label={t('appPageBreakPreview')} symbol="┆" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupShow')}>
          <div className="check-column">
            <button className="check-item" onClick={() => onCommand('toggle-gridlines')}>
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appGridlines')}
            </button>
            <span className="check-item reserved-check">
              <i className="check-box">✓</i>
              {t('appFormulaBar')}
            </span>
            <span className="check-item reserved-check">
              <i className="check-box">✓</i>
              {t('appHeadings')}
            </span>
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
          <RibbonReserved large label={t('appZoomToSelection')} symbol="⌖" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupWindow')}>
          {largeMenu(t('appFreezePanes'), '❄', t('appFreezeTitle'), [
            { value: 'freeze-here', label: t('appFreezeAtSelection') },
            { value: 'freeze-top-row', label: t('appFreezeTopRow') },
            { value: 'freeze-first-col', label: t('appFreezeFirstCol') },
            { value: 'unfreeze', label: t('appUnfreeze') },
          ])}
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="🗔" />
              {t('appNewWindow')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="▦" />
              {t('appArrangeAll')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="◫" />
              {t('appSplit')}
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appMacros')}>
          <RibbonReserved large menu label={t('appMacros')} symbol="▶" />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Review') {
    return (
      <div className="ribbon">
        <RibbonGroup label={t('appGroupProofing')}>
          <RibbonReserved large label={t('appSpelling')} symbol="abc" />
          <RibbonReserved large label={t('appThesaurus')} symbol="🕮" />
          <RibbonButton
            large
            label={t('appWorkbookStatsLabel')}
            detail={t('appSheetsCellsFormulas')}
            symbol="🧮"
            onClick={() => onCommand('workbook-statistics')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupAccessibility')}>
          <RibbonReserved large menu label={t('appCheckAccessibility')} symbol="♿" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupInsights')}>
          <RibbonReserved large label={t('appSmartLookup')} symbol="💡" />
          <RibbonReserved large label={t('appTranslate')} symbol="文" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupComments')}>
          <RibbonButton
            large
            label={t('appNewComment')}
            detail={t('appNewCommentDetail')}
            symbol="🗨"
            onClick={() => onCommand('note-open')}
          />
          <RibbonButton
            large
            label={t('appDeleteLabel')}
            detail={t('appNoteAtSelection')}
            symbol="🗑"
            onClick={() => onCommand('note-delete')}
          />
          <div className="row-stack">
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="←" />
              {t('appPrevious')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="→" />
              {t('appNext')}
            </span>
            <span className="styles-row reserved" title={t('appNotAvailableYet')}>
              <ToolSymbol symbol="🗨" />
              {t('appShowComments')}
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupNotes')}>
          <RibbonButton
            large
            menu
            label={t('appGroupNotes')}
            detail={t('appNotesDetail')}
            symbol="🗨"
            onClick={() => onCommand('note-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupProtection')}>
          <RibbonButton
            large
            label={t(sheetProtected ? 'appUnprotectSheet' : 'appProtectSheet')}
            detail={t(sheetProtected === null ? 'appOpenFileFirst' : 'appNoPassword')}
            symbol={sheetProtected ? '🔓' : '🔒'}
            onClick={() => onCommand('sheet-protect')}
          />
          <RibbonReserved large label={t('appProtectWorkbook')} symbol="🔐" />
          <RibbonReserved large label={t('appAllowEditRanges')} symbol="⬚" />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupInk')}>
          <RibbonReserved large menu label={t('appHideInk')} symbol="✒" />
        </RibbonGroup>
      </div>
    )
  }

  const fontFamilies = ['Aptos', 'Arial', 'Calibri', 'Times New Roman', '微软雅黑', '宋体']
  const fontSizes = [9, 10, 11, 12, 14, 16, 18, 22, 26]
  const echoFamily = selectionFormat?.fontFamily ?? 'Aptos'
  const echoSize = selectionFormat?.fontSize ?? 11
  const familyOptions = fontFamilies.includes(echoFamily)
    ? fontFamilies
    : [echoFamily, ...fontFamilies]
  const sizeOptions = fontSizes.includes(echoSize)
    ? fontSizes
    : [...fontSizes, echoSize].sort((a, b) => a - b)
  return (
    <div className="ribbon">
      <RibbonGroup label={t('appGroupAiAssistant')}>
        <button
          className={`ribbon-tool as-button large ai-entry ${aiOpen ? 'active' : ''}`}
          title={t('aiOpenAssistant')}
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
          title={t('aiCheckPrompt')}
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
                <circle cx="10.5" cy="10.5" r="6.2" />
                <path d="m15.3 15.3 5.2 5.2" strokeLinecap="round" />
                <path d="m7.9 10.7 1.9 1.9 3.3-3.3" strokeLinecap="round" strokeLinejoin="round" />
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
          title={t('aiAnalyzePrompt')}
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
                <path d="M4.5 4.5v15h15" strokeLinecap="round" />
                <path d="M9 15.5v-4M13 15.5V7.5M17 15.5v-6" strokeLinecap="round" />
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
          title={t('appPasteTitle')}
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
        ])}
        <div className="tool-stack">
          <button title={t('appCutTitle')} onClick={() => onCommand('cut')}>
            <ToolSymbol symbol="✂" />
          </button>
          <button title={t('appCopyTitle')} onClick={() => onCommand('copy')}>
            <ToolSymbol symbol="⧉" />
          </button>
          <button title={t('appFormatPainter')} onClick={() => onCommand('format-painter')}>
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
              value={echoFamily}
              options={familyOptions.map((family) => ({ value: family, label: family }))}
              onPick={(value) => onCommand(`font-family:${value}`)}
              commit={(text) => onCommand(`font-family:${text}`)}
            />
            <EditableMenuSelect
              className="select-like font-size"
              label="Font size"
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
              title={t('appIncreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, 1)}`)}
            >
              <ToolSymbol symbol="A↑" />
            </button>
            <button
              title={t('appDecreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, -1)}`)}
            >
              <ToolSymbol symbol="A↓" />
            </button>
          </div>
          <div className="inline-tools">
            <button
              title={t('appBold')}
              className={selectionFormat?.bold ? 'is-active' : ''}
              onClick={() => onCommand('bold')}
            >
              <b>B</b>
            </button>
            <button
              title={t('appItalic')}
              className={selectionFormat?.italic ? 'is-active' : ''}
              onClick={() => onCommand('italic')}
            >
              <em>I</em>
            </button>
            <button
              title={t('appUnderline')}
              className={selectionFormat?.underline ? 'is-active' : ''}
              onClick={() => onCommand('underline')}
            >
              <u>U</u>
            </button>
            <button title={t('appDoubleUnderline')} onClick={() => onCommand('underline:double')}>
              <u style={{ textDecorationStyle: 'double' }}>D</u>
            </button>
            <button
              title={t('appStrikethrough')}
              className={selectionFormat?.strike ? 'is-active' : ''}
              onClick={() => onCommand('strike')}
            >
              <s>S</s>
            </button>
            <label className="color-tool" title={t('appFontColor')}>
              <span className="swatch-letter">
                A<i style={{ background: fontColor }} />
              </span>
              <input
                type="color"
                aria-label="Font color"
                value={fontColor}
                onChange={(event) => {
                  setFontColor(event.target.value)
                  onCommand(`font-color:${event.target.value}`)
                }}
              />
            </label>
            <label className="color-tool" title={t('appFillColor')}>
              <span className="swatch-letter">
                <ToolSymbol symbol="◧" />
                <i style={{ background: fillColor }} />
              </span>
              <input
                type="color"
                aria-label="Fill color"
                value={fillColor}
                onChange={(event) => {
                  setFillColor(event.target.value)
                  onCommand(`fill:${event.target.value}`)
                }}
              />
            </label>
            <button title={t('dlgFcNoFill')} onClick={() => onCommand('fill:none')}>
              <ToolSymbol symbol="∅" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Borders"
              title={t('appBorders')}
              display={
                <>
                  <ToolSymbol symbol="⊡" /> {t('appBorders')}
                </>
              }
              options={[
                { value: 'all', label: t('appBorderAll') },
                { value: 'outer', label: t('appBorderOuter') },
                { value: 'thick-outer', label: t('appBorderThickOuter') },
                { value: 'top', label: t('appBorderTop') },
                { value: 'bottom', label: t('appBorderBottom') },
                { value: 'left', label: t('appBorderLeft') },
                { value: 'right', label: t('appBorderRight') },
                { value: 'none', label: t('appBorderNone') },
              ]}
              onPick={(value) => onCommand(`border:${value}:${borderColor}`)}
            />
            <label className="color-tool" title={t('appBorderColor')}>
              <span className="swatch-letter">
                <ToolSymbol symbol="⊡" />
                <i style={{ background: borderColor }} />
              </span>
              <input
                type="color"
                aria-label="Border color"
                value={borderColor}
                onChange={(event) => setBorderColor(event.target.value)}
              />
            </label>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupAlignment')}>
        <div className="ribbon-rows">
          <div className="inline-tools alignment-tools">
            <button title={t('appTopAlign')} onClick={() => onCommand('valign:top')}>
              <ToolSymbol symbol="⤒" />
            </button>
            <button title={t('appMiddleAlign')} onClick={() => onCommand('valign:middle')}>
              <ToolSymbol symbol="↕" />
            </button>
            <button title={t('appBottomAlign')} onClick={() => onCommand('valign:bottom')}>
              <ToolSymbol symbol="⤓" />
            </button>
            <button
              title={t('dlgFcWrapText')}
              className={selectionFormat?.wrap ? 'is-active' : ''}
              onClick={() => onCommand('wrap')}
            >
              <ToolSymbol symbol="↩" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Orientation"
              title={t('dlgFcOrientation')}
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
            <button title={t('appAlignLeft')} onClick={() => onCommand('align:left')}>
              <ToolSymbol symbol="≡" />
            </button>
            <button title={t('appAlignCenter')} onClick={() => onCommand('align:center')}>
              <ToolSymbol symbol="≣" />
            </button>
            <button title={t('appAlignRight')} onClick={() => onCommand('align:right')}>
              <ToolSymbol symbol="☰" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Merge cells"
              title={t('appMergeCells')}
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
            <button title={t('appCurrency')} onClick={() => onCommand('format:$#,##0.00')}>
              $
            </button>
            <button title={t('appPercentTitle')} onClick={() => onCommand('format:0.00%')}>
              %
            </button>
            <button title={t('appThousands')} onClick={() => onCommand('format:#,##0')}>
              ,
            </button>
            <button title={t('appIncreaseDecimal')} onClick={() => onCommand('decimal-inc')}>
              .0+
            </button>
            <button title={t('appDecreaseDecimal')} onClick={() => onCommand('decimal-dec')}>
              .0−
            </button>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupStyles')}>
        <div className="styles-stack">
          <button className="styles-row as-button" onClick={() => onCommand('cf-open')}>
            <ToolSymbol symbol="▤" />
            {t('appConditionalFormatting')}
            <CaretIcon />
          </button>
          <div className="styles-row as-button" title={t('appFormatAsTableTitle')}>
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
          <div className="styles-row as-button" title={t('appCellStylesTitle')}>
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
          <button className="styles-row as-button" onClick={() => onCommand('format-cells')}>
            <ToolSymbol symbol="🎨" />
            {t('appFormatCells')} ⌘1
            <CaretIcon />
          </button>
          <div className="inline-tools cell-tools">
            <button title={t('appInsertRow')} onClick={() => onCommand('insert-row-here')}>
              <ToolSymbol symbol="⤒" />
            </button>
            <button title={t('appDeleteRow')} onClick={() => onCommand('delete-row-here')}>
              <ToolSymbol symbol="⤓" />
            </button>
            <button title={t('appInsertCol')} onClick={() => onCommand('insert-col-here')}>
              <ToolSymbol symbol="⇤" />
            </button>
            <button title={t('appDeleteCol')} onClick={() => onCommand('delete-col-here')}>
              <ToolSymbol symbol="⇥" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Format"
              title={t('appFormatMenu')}
              display={
                <>
                  <ToolSymbol symbol="⇳" /> {t('appFormatMenu')}
                </>
              }
              options={[
                { value: 'row-height-open', label: `${t('appRowHeight')}…` },
                { value: 'col-width-open', label: `${t('appColWidth')}…` },
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
              label="Fill"
              title={t('appFillMenu')}
              display={
                <>
                  <ToolSymbol symbol="↓" /> {t('appFillMenu')}
                </>
              }
              options={[
                { value: 'down', label: t('appFillDown') },
                { value: 'right', label: t('appFillRight') },
              ]}
              onPick={(value) => onCommand(`fill-${value}`)}
            />
            <MenuSelect
              className="select-like compact"
              label="Clear"
              title={t('appClear')}
              display={
                <>
                  <ToolSymbol symbol="⌫" /> {t('appClear')}
                </>
              }
              options={[
                { value: 'all', label: t('appClearAll') },
                { value: 'formats', label: t('appClearFormats') },
                { value: 'contents', label: t('appClearContents') },
              ]}
              onPick={(value) => onCommand(`clear-${value}`)}
            />
          </div>
          <div className="inline-tools">
            <button className="labeled" title={t('appFindTitle')} onClick={() => onCommand('find')}>
              <ToolSymbol symbol="🔍" /> {t('appFind')}
            </button>
            <button
              className="labeled"
              title={t('appReplace')}
              onClick={() => onCommand('replace')}
            >
              <ToolSymbol symbol="⇄" /> {t('appReplace')}
            </button>
            <button
              className="labeled"
              title={`${t('appGoTo')} ⌘G`}
              onClick={() => onCommand('goto-open')}
            >
              <ToolSymbol symbol="⌖" /> {t('appGoTo')}
            </button>
          </div>
        </div>
      </RibbonGroup>
      <div className="ribbon-genspark-sep" aria-hidden />
      <button
        className="ribbon-genspark-btn"
        title={t('aiOpenAssistant')}
        onClick={() => onCommand('ai-toggle-panel')}
      >
        <GensparkMark size={28} />
        <span>Genspark</span>
      </button>
    </div>
  )
}

/// Custom ribbon dropdown replacing the native <select>: macOS pops the native
/// menu over the control (covering ribbon content), while this panel is
/// anchored below its trigger — same pattern as the slides ribbon's .rb-drop.
function MenuSelect({
  label,
  title,
  className,
  cover = false,
  display,
  value = '',
  options,
  onPick,
}: {
  /// aria-label (mirrors the old select's aria-label)
  readonly label: string
  readonly title?: string
  /// trigger classes for the combobox look (e.g. 'select-like compact')
  readonly className?: string
  /// invisible trigger stretched over the host tool (old .cover-select)
  readonly cover?: boolean
  /// content shown in the trigger (combobox variants)
  readonly display?: React.ReactNode
  /// currently applied value, highlighted in the open panel ('' = none)
  readonly value?: string
  readonly options: readonly { value: string; label: string }[]
  readonly onPick: (value: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const onDown = (event: MouseEvent): void => {
      if (!wrapRef.current?.contains(event.target as Node)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])
  return (
    <div ref={wrapRef} className={`menu-select${cover ? ' menu-select-cover' : ''}`}>
      <button
        type="button"
        className={cover ? 'cover-select' : className}
        title={title}
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
              {option.label}
            </button>
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
  className,
  value,
  options,
  onPick,
  commit,
}: {
  readonly label: string
  readonly className?: string
  readonly value: string
  readonly options: readonly { value: string; label: string }[]
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
  useEffect(() => {
    if (!open) return
    const onDown = (event: MouseEvent): void => {
      if (!wrapRef.current?.contains(event.target as Node)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])
  const commitDraft = (): void => {
    const text = draftRef.current?.trim()
    if (text && text !== value) commit(text)
    setDraft(null)
  }
  return (
    <div ref={wrapRef} className="menu-select">
      <span className={`${className ?? ''} menu-select-edit`}>
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
          onClick={() => setOpen((v) => !v)}
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
              className={option.value === value ? 'on' : ''}
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

// display only — option values keep the English label as the command identity
const NUMBER_FORMAT_LABEL: Record<string, StringKey> = {
  General: 'dlgFcNumGeneral',
  Number: 'dlgFcNumNumber',
  Currency: 'dlgFcNumCurrency',
  Accounting: 'appNumFmtAccounting',
  'Short Date': 'appNumFmtShortDate',
  'Long Date': 'appNumFmtLongDate',
  Time: 'dlgFcNumTime',
  Percentage: 'dlgFcNumPercent',
  Fraction: 'appNumFmtFraction',
  Scientific: 'dlgFcNumScientific',
  Text: 'dlgFcNumText',
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
      title={pattern || t('dlgFcNumGeneral')}
      value={current}
      display={localized(current)}
      options={NUMBER_FORMAT_CATEGORIES.map((category) => ({
        value: category.label,
        label: localized(category.label),
      }))}
      onPick={(value) => {
        const category = NUMBER_FORMAT_CATEGORIES.find((candidate) => candidate.label === value)
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
      title={`${label} — ${detail}`}
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

/// A greyed placeholder item: reserves its ribbon slot, not built yet.
function RibbonReserved({
  label,
  symbol,
  menu = false,
  large = false,
}: {
  readonly label: string
  readonly symbol: string
  readonly menu?: boolean
  readonly large?: boolean
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <div className={`ribbon-tool reserved ${large ? 'large' : ''}`} title={t('appNotAvailableYet')}>
      {large ? (
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
        <small>{t('appComingSoon')}</small>
      </span>
    </div>
  )
}
