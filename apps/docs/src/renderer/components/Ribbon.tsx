import { memo, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { Lang } from '@genoffice/i18n'
import type {
  ChangeEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
} from 'react'
import type { ChainedCommands, Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import type { Mark, Node as PMNode, ResolvedPos } from '@tiptap/pm/model'
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  selectedRect,
  setCellAttr,
} from '@tiptap/pm/tables'
import type {
  Block,
  CustomNumberingLevel,
  DocDefaults,
  HeaderFooter,
  Run,
  SectionSettings,
  SourceInfo,
  StyleInfo,
  TableAutoFitMode,
  TableLook,
  TextboxDisplay,
  TextboxParaDisplay,
  ThemeColors,
  ThemeFonts,
} from '@genoffice/docx-engine'
import {
  Dropdown,
  isSymbolFontFamily,
  useDismissablePopover,
  useRibbonCollapse,
} from '@genoffice/ui'
import { HIGHLIGHT_CSS } from '../editor/extensions'
import { applyCase, type CaseMode } from '../editor/case-transform'
import { isRtlUiLang, setParagraphDirection, setSelectionAlign } from '../editor/direction'
import { setInactiveSelectionShown } from '../editor/inactive-selection'
import { stepParagraphIndent } from '../editor/indent'
import { pasteFromClipboard } from '../editor/paste-actions'
import type { PasteMode } from '../editor/paste-options'
import { formatNumber } from '../editor/numbering'
import { getActiveSubEditor } from '../editor/active-editor'
import { wordUnitAt } from '../editor/word-range'
import {
  BULLET_LIBRARY,
  MULTILEVEL_LIBRARY,
  NUMBER_LIBRARY,
  bulletPresetLevels,
  numberPresetLevels,
  previewLevelText,
  rememberListPreset,
  type ListPresetKind,
} from '../list-presets'
import type { ListDialogKind } from './ListDialogs'
import type { InkTool } from '../editor/ink'
import type { RibbonFormatState } from './ribbon-format-state'
import { distributeSelectedColumns, setSelectedColumnWidth } from '../editor/table-sizing'
import {
  distributeRowsEvenly,
  enterSelectedTable,
  selectTablePart,
  tableCellsSelection,
  setCellAlignment,
  setCellTextDirection,
  splitTableAtSelection,
  type CellHAlign,
  type CellTextDirection,
  type CellVAlign,
  type TableSelectKind,
} from '../editor/table-ops'
import { TABLE_AUTO_FIT_OPTIONS, type TablePropertiesTab } from './TablePropertiesDialog'
import type { TableDialogKind } from './TableDialogs'
import {
  IconCellAlign,
  IconCellMargins,
  IconDeleteCells,
  IconDistributeColumns,
  IconDistributeRows,
  IconGridlines,
  IconInsertCells,
  IconSelectCells,
  IconSplitTable,
  IconTextDirection,
} from './table-icons'
import {
  TABLE_BORDER_MENU,
  setSelectionBorders,
  type TableBorderMode,
} from '../editor/table-borders'
import {
  applyTablePreset,
  repeatHeaderState,
  setTableAutoFit,
  setTableLookOption,
  toggleRepeatHeaderRows,
} from '../editor/table-properties'
import { useI18n, type StringKey } from '../i18n/locale'
import { pxToTwips, twipsToPx } from '../units'
import { LengthInput } from './LengthInput'
import {
  FALLBACK_STYLES,
  activeStyleKey as activeStyleKeyOf,
  collectUsedStyleIds,
  quickStyleEntries,
  styleKeyOf,
  styleLabel,
  stylePreviewCss,
} from '../style-gallery'
import { RibbonColorPalette } from './ribbon-color-palette'
import { HeaderFooterTab, type HfAction, type HfEditingInfo } from './ribbon-hf-tab'
import { PasteSpecialDialog } from './PasteSpecialDialog'
import { fontFamiliesFor, isEastAsianFontName } from '../font-list'
import {
  fontSizeLabel,
  fontSizeOptions,
  parseFontSize,
  stepFontSize as stepSizeInList,
} from '../font-sizes'
import { useSystemFontFamilies } from '../system-fonts'
import { cssFontFamily } from '../line-metrics'
import {
  DesignTab,
  DrawTab,
  imageSizeOf,
  InsertTab,
  LayoutTab,
  ReferencesTab,
  ReviewTab,
  ViewTab,
  type InkPenSettings,
  type RevisionDisplayMode,
  type ViewMode,
  applyGalleryStyle,
  setParaAttrs,
  toggleParaSpace as toggleParaSpaceImpl,
} from './ribbon-tabs'
import { WRAP_OPTIONS } from './ContextMenu'
import { CropDialog, CutoutDialog } from './PictureDialogs'
import {
  GensparkMark,
  IconAlignCenter,
  IconAlignJustify,
  IconAlignLeft,
  IconAlignRight,
  IconAutoFit,
  IconBorderAll,
  IconBorderBottom,
  IconBorderInner,
  IconBorderInsideH,
  IconBorderInsideV,
  IconBorderLeft,
  IconBorderNone,
  IconBorderOuter,
  IconBorderRight,
  IconBorderTop,
  IconBullets,
  IconCaret,
  IconColDelete,
  IconColInsertLeft,
  IconColInsertRight,
  IconMultilevel,
  IconClearFormat,
  IconCopy,
  IconCut,
  IconFormatPainter,
  IconGrowFont,
  IconHighlight,
  IconIndentDec,
  IconIndentInc,
  IconCrop,
  IconDirLtr,
  IconDirRtl,
  IconLineSpacing,
  IconMergeCells,
  IconNumbered,
  IconPaste,
  IconPasteMerge,
  IconPasteSource,
  IconPasteText,
  IconPilcrow,
  IconFlipH,
  IconFlipV,
  IconRemoveBg,
  IconReplacePicture,
  IconRotateLeft,
  IconRotateRight,
  IconRowDelete,
  IconRowInsertAbove,
  IconRowInsertBelow,
  IconShading,
  IconChangeCase,
  IconFontColorA,
  IconShrinkFont,
  IconSplitCells,
  IconSubscript,
  IconSuperscript,
  IconTableDelete,
  IconRepeatHeader,
  IconTableProperties,
  IconReplace,
  IconSearch,
  IconStylesPane,
  IconSelectAll,
} from './icons'
interface RibbonProps {
  /** App keyboard shortcuts reuse ribbon closures through here (font-size stepping keeps its coalescing) */
  actionsRef?: React.MutableRefObject<{
    stepFontSize?: (dir: 1 | -1) => void
    nudgeFontSize?: (dir: 1 | -1) => void
  }>
  /** Quick-access area on the tab row's left (save/undo-redo/autosave), matching the WPS/Office QAT */
  quickActions?: React.ReactNode
  /** Right side of the tab row (file name, etc.) */
  trailingActions?: React.ReactNode
  editor: Editor
  /** shallow-stable snapshot of every editor-state read shown in the ribbon (memo invalidation key) */
  formatState: RibbonFormatState
  hasDoc: boolean
  blocks: Block[]
  /** Fallback when a new list can't reuse a numId (adopt a document definition / create one) */
  allocateNumId?: (kind: 'bullet' | 'ordered') => string | null
  /** New list definitions with custom levels (bullet library / numbering library / multilevel list) */
  createListDef?: (levels: CustomNumberingLevel[]) => string | null
  /** the host owns the list dialogs (define bullet / number / multilevel, numbering value, indents) */
  onListDialog?: (kind: ListDialogKind) => void
  /** formats already used in the document, for the "Document …" gallery sections */
  documentListPresets?: Record<ListPresetKind, CustomNumberingLevel[][]>
  /** recently picked presets (localStorage), refreshed by the host after each pick */
  recentListPresets?: Record<ListPresetKind, CustomNumberingLevel[][]>
  /** document character styles, from ParsedDoc.styles (type === 'character') */
  styles?: Map<string, StyleInfo>
  /** document-wide text defaults from styles.xml */
  docDefaults?: DocDefaults
  /** Open the paragraph dialog (line-spacing rule / exact value entry lives there) */
  onParagraphDialog?: () => void
  /** Paste ▸ Set Default Paste…: the host's preferences dialog holds the default paste mode */
  onPasteDefaults?: () => void
  /** Word's table dialogs live in App so the right-click menu and the Table menu share them */
  onTableDialog?: (kind: TableDialogKind, tab?: TablePropertiesTab) => void
  /** screen-only dashed gridlines on borderless tables (Table Design ▸ View Gridlines) */
  tableGridlines?: boolean
  onToggleTableGridlines?: () => void
  /** Home ▸ Editing: open the Find panel on its Find / Replace tab */
  onFind?: () => void
  onReplace?: () => void
  onOpen: () => void
  onSave: () => void
  onSaveAs: () => void
  showAi: boolean
  onToggleAi: () => void
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  /** Multi-section documents: index of the cursor's section (0-based); null for single-section */
  activeSection: number | null
  onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => void
  /** paper size for every section (More Paper Sizes… > Whole document); portrait dimensions */
  onPaperSizeAll: (
    portraitW: number,
    portraitH: number,
    orientation: SectionSettings['orientation'],
  ) => void
  mirrorMargins: boolean
  onMirrorMargins: (on: boolean) => void
  pageColor: string | null
  onPageColor: (hex: string | null) => void
  /** Design → Watermark / Themes */
  watermark: string | null
  onWatermark: (text: string | null) => void
  themeFonts: ThemeFonts | null
  onThemeFonts: (fonts: ThemeFonts) => void
  themeColors: ThemeColors | null
  onThemeColors: (colors: ThemeColors) => void
  /** Draw → pen / highlighter / eraser */
  inkTool: InkTool
  onInkTool: (tool: InkTool) => void
  inkPen: InkPenSettings
  onInkPen: (settings: InkPenSettings) => void
  inkHighlighter: InkPenSettings
  onInkHighlighter: (settings: InkPenSettings) => void
  inkCount: number
  onInkClearAll: () => void
  /** References → footnotes / endnotes / citations */
  onInsertNote: (kind: 'footnote' | 'endnote') => void
  sources: SourceInfo[]
  /** footnotes/endnotes hold Zotero citation fields the bridge cannot see yet */
  zoteroNoteFields?: boolean
  onAddSource: (source: SourceInfo) => void
  /** TOC page-number backfill: docHeadings in document order → real page numbers (null when not computable) */
  headingPages?: () => number[] | null
  zoom: number
  onZoom: (zoom: number) => void
  /** compute zoom from the current window size (Word: page width / whole page) */
  onZoomFit: (mode: 'width' | 'page') => void
  onZoomDialog: () => void
  darkPage: boolean
  onDarkPage: (v: boolean) => void
  onAiPreset: (instruction: string) => void
  /** external request (e.g. native menu Page Setup) to switch to a specific tab */
  tabRequest?: { tab: string; nonce: number } | null
  header: HeaderFooter | null
  onHeader: (next: HeaderFooter) => void
  onPageNumFormat: () => void
  onInsertField: (instr: string) => void
  footer: HeaderFooter | null
  onFooter: (next: HeaderFooter) => void
  /** a header/footer strip is being edited: the contextual Header & Footer tab shows */
  hfEditing: HfEditingInfo | null
  onHfAction: (action: HfAction) => void
  /** Insert → Header/Footer → Edit: open the strip's editor */
  onHfEdit: (kind: 'header' | 'footer') => void
  showMarks: boolean
  onShowMarks: (v: boolean) => void
  showRuler: boolean
  onShowRuler: (v: boolean) => void
  showNav: boolean
  onShowNav: (v: boolean) => void
  /** Home > Styles > Styles Pane (Word for Mac) */
  showStylesPane?: boolean
  onShowStylesPane?: (v: boolean) => void
  commentCount: number
  /** unresolved root comments (drives the AI resolve-comments action) */
  openCommentCount: number
  resolvedCommentCount: number
  onShowComments: () => void
  /** Review → comments / revisions / compare / protection */
  canComment: boolean
  onNewComment: () => void
  commentAtCaret: boolean
  onDeleteComment: () => void
  onDeleteAllComments: (resolvedOnly: boolean) => void
  onGotoComment: (dir: 1 | -1) => void
  trackChanges: boolean
  onTrackChanges: (on: boolean) => void
  /** native check-as-you-type spellcheck (red squiggle) */
  spellcheck: boolean
  onSpellcheck: (on: boolean) => void
  revisionDisplay: RevisionDisplayMode
  onRevisionDisplay: (mode: RevisionDisplayMode) => void
  revisionCount: number
  onAcceptRevision: (all: boolean) => void
  onRejectRevision: (all: boolean) => void
  onGotoRevision: (dir: 1 | -1) => void
  isProtected: boolean
  /** comments restriction: adding comments stays allowed although the body is read-only */
  commentsAllowed: boolean
  /** trackedChanges restriction: the recorder is forced on (toggle and accept/reject disabled) */
  trackChangesForced: boolean
  /** any protection is configured (highlights the Protect Document button) */
  protectActive: boolean
  onProtectDoc: () => void
  onCompare: () => void
  /** current document path (View → New Window opens it in another window) */
  filePath: string | null
  viewMode: ViewMode
  onViewMode: (mode: ViewMode) => void
  readMode: boolean
  onReadMode: (v: boolean) => void
  showGrid: boolean
  onShowGrid: (v: boolean) => void
  splitView: boolean
  onSplitView: (v: boolean) => void
  onPagePreview: () => void
}

export interface PainterFormat {
  marks: Array<{ type: string; attrs: Record<string, unknown> }>
  /** source paragraph's node type + formatting attrs (null when the caret is not in a paintable block) */
  block: { type: string; attrs: Record<string, unknown> } | null
}

interface PainterState extends PainterFormat {
  /** double-click on the button: stays armed until Esc or another click */
  locked: boolean
}

/** Character-formatting marks the painter transfers; semantic marks (links,
 *  comments, revisions, fields) are neither picked up nor stripped from the target. */
const PAINTER_MARK_TYPES = ['bold', 'italic', 'underline', 'strike', 'docTextStyle']

/** Paragraph-formatting attrs the painter transfers. Identity/anchor attrs
 *  (docxIndex, bookmarks, comment ranges, revisions, sdtShell…) stay with the target. */
const PAINTER_PARA_KEYS = [
  'styleId',
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'snapToGrid',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'bidi',
  'autoSpace',
  'shadingFill',
  'emptyRunSize',
  'borders',
  'borderLines',
  'tabStops',
]

/** Per-block-type attrs that define the block's identity as formatting (heading level, list numbering) */
const PAINTER_BLOCK_EXTRA: Record<string, string[]> = {
  docParagraph: [],
  docHeading: ['level'],
  docListItem: ['kind', 'numId', 'ilvl'],
}

/** the word under a painter click: Word brushes the word and its trailing spaces */
function painterWordRangeAt($pos: ResolvedPos): { from: number; to: number } | null {
  const para = $pos.parent
  if (!para.isTextblock) return null
  const text = para.textBetween(0, para.content.size, '\0', '\0')
  const base = $pos.start()
  const span = wordUnitAt(text, $pos.parentOffset)
  // no word under the pointer: nothing to mark, the paragraph format still applies
  if (!span) return { from: base + $pos.parentOffset, to: base + $pos.parentOffset }
  return { from: base + span.start, to: base + span.end }
}

export function applyPainterFormat(
  editor: Editor,
  fmt: PainterFormat,
  from: number,
  to: number,
  caretAfter: number | null,
): void {
  // Word: every brush stroke is its own undo step, even quick strokes on
  // neighbouring paragraphs that history would otherwise merge into one
  let c = editor
    .chain()
    .focus()
    .command(({ tr }) => {
      closeHistory(tr)
      return true
    })
    .setTextSelection({ from, to })
  if (to > from) {
    // strip only formatting marks, then re-add the picked-up ones: semantic
    // marks on the target (links, comments, revisions) survive the brush
    for (const t of PAINTER_MARK_TYPES) c = c.unsetMark(t)
    for (const m of fmt.marks) c = c.setMark(m.type, m.attrs)
  }
  c = c.command(({ tr }) => {
    const block = fmt.block
    if (!block) return true
    const type = editor.schema.nodes[block.type]
    if (!type) return true
    const sel = tr.selection
    const jobs: Array<{ pos: number; attrs: Record<string, unknown> }> = []
    tr.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
      if (!(node.type.name in PAINTER_BLOCK_EXTRA)) return true
      // keep the target's identity attrs, overwrite every formatting attr
      // (explicit nulls in block.attrs reset what the source didn't set)
      jobs.push({ pos, attrs: { ...node.attrs, ...block.attrs } })
      return false
    })
    for (const job of jobs) tr.setNodeMarkup(job.pos, type, job.attrs)
    return true
  })
  if (caretAfter != null) c = c.setTextSelection(caretAfter)
  c.run()
}

// Word for Mac has no File ribbon tab: file actions live in the native menu
// bar (which we provide). Windows Word does have one, so keep it there.
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
// Word offers Half-width / Full-width under Change Case only on CJK UIs
const WIDTH_CASE_LANGS = new Set<Lang>(['zh', 'zh-TW', 'ja'])
/** shell tab mode: the tab strip above owns traffic lights / caption buttons */
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'

const TABS = (
  IS_MAC
    ? ['home', 'insert', 'draw', 'design', 'layout', 'references', 'review', 'view']
    : ['file', 'home', 'insert', 'draw', 'design', 'layout', 'references', 'review', 'view']
) as readonly string[]
const TABLE_TABS = ['tableDesign', 'tableLayout'] as const

/** OOXML ST_Border values the renderer draws distinctly; glyph labels keep the picker language-neutral */
const BORDER_STYLE_OPTIONS = [
  { value: 'single', label: '\u2500\u2500\u2500\u2500\u2500' },
  { value: 'thick', label: '\u2501\u2501\u2501\u2501\u2501' },
  { value: 'double', label: '\u2550\u2550\u2550\u2550\u2550' },
  { value: 'triple', label: '\u2261\u2261\u2261\u2261\u2261' },
  { value: 'dotted', label: '\u2504\u2504\u2504\u2504\u2504' },
  { value: 'dashed', label: '\u254c\u254c\u254c\u254c\u254c' },
  { value: 'dotDash', label: '\u2500\u00b7\u2500\u00b7\u2500' },
  { value: 'dotDotDash', label: '\u2500\u00b7\u00b7\u2500\u00b7\u00b7' },
] as const

/** Borders ▾ entries: label + gallery glyph, in TABLE_BORDER_MENU (Word) order */
const TABLE_BORDER_ITEMS: Record<
  TableBorderMode,
  { label: StringKey; Icon: (props: { size?: number }) => ReactNode }
> = {
  bottom: { label: 'ribbonBorderBottom', Icon: IconBorderBottom },
  top: { label: 'ribbonBorderTop', Icon: IconBorderTop },
  left: { label: 'ribbonBorderLeft', Icon: IconBorderLeft },
  right: { label: 'ribbonBorderRight', Icon: IconBorderRight },
  none: { label: 'ribbonNoBorders', Icon: IconBorderNone },
  all: { label: 'ribbonAllBorders', Icon: IconBorderAll },
  outer: { label: 'ribbonOuterBorders', Icon: IconBorderOuter },
  inner: { label: 'ribbonInnerBorders', Icon: IconBorderInner },
  insideH: { label: 'ribbonTableInsideHBorders', Icon: IconBorderInsideH },
  insideV: { label: 'ribbonTableInsideVBorders', Icon: IconBorderInsideV },
}

/** Table Layout ▸ Select ▾ in Word for Mac order */
const TABLE_SELECT_ITEMS: Array<[TableSelectKind, StringKey]> = [
  ['cell', 'ribbonSelectCell'],
  ['column', 'ribbonSelectColumn'],
  ['row', 'ribbonSelectRow'],
  ['table', 'ribbonSelectTable'],
]

const CELL_ALIGN_GRID: Array<[CellVAlign, CellHAlign, StringKey]> = [
  ['top', 'left', 'ribbonAlignTopLeft'],
  ['top', 'center', 'ribbonAlignTopCenter'],
  ['top', 'right', 'ribbonAlignTopRight'],
  ['center', 'left', 'ribbonAlignMiddleLeft'],
  ['center', 'center', 'ribbonAlignMiddleCenter'],
  ['center', 'right', 'ribbonAlignMiddleRight'],
  ['bottom', 'left', 'ribbonAlignBottomLeft'],
  ['bottom', 'center', 'ribbonAlignBottomCenter'],
  ['bottom', 'right', 'ribbonAlignBottomRight'],
]

const CELL_TEXT_DIRECTIONS: Array<[CellTextDirection, StringKey]> = [
  ['lrTb', 'ribbonTextDirectionHorizontal'],
  ['tbRl', 'ribbonTextDirectionRotate90'],
  ['btLr', 'ribbonTextDirectionRotate270'],
]
const IMAGE_TABS = ['pictureFormat'] as const
const SHAPE_TABS = ['shapeFormat'] as const
const HF_TABS = ['headerFooter'] as const
type RibbonTab =
  | (typeof TABS)[number]
  | (typeof TABLE_TABS)[number]
  | (typeof IMAGE_TABS)[number]
  | (typeof SHAPE_TABS)[number]
  | (typeof HF_TABS)[number]

const TABLE_PRESETS = [
  {
    label: 'ribbonTablePresetGrid',
    headerFill: null,
    band1Fill: null,
    band2Fill: null,
    borderColor: '808080',
  },
  {
    label: 'ribbonTablePresetBlueHeader',
    headerFill: 'D9EAF7',
    band1Fill: null,
    band2Fill: null,
    borderColor: '5B9BD5',
  },
  {
    label: 'ribbonTablePresetBlueBanded',
    headerFill: 'BDD7EE',
    band1Fill: 'DDEBF7',
    band2Fill: 'FFFFFF',
    borderColor: '9DC3E6',
  },
  {
    label: 'ribbonTablePresetGrayBanded',
    headerFill: 'D9E1F2',
    band1Fill: 'E7E6E6',
    band2Fill: 'FFFFFF',
    borderColor: 'A5A5A5',
  },
  {
    label: 'ribbonTablePresetGreenHeader',
    headerFill: 'E2F0D9',
    band1Fill: null,
    band2Fill: null,
    borderColor: '70AD47',
  },
] as const satisfies ReadonlyArray<{
  label: StringKey
  headerFill: string | null
  band1Fill: string | null
  band2Fill: string | null
  borderColor: string
}>

// tab values double as internal-state / external tabRequest keys; translated for display via these string keys
const TAB_LABEL_KEYS: Record<string, StringKey> = {
  file: 'ribbonTabFile',
  home: 'ribbonTabHome',
  insert: 'ribbonTabInsert',
  draw: 'ribbonTabDraw',
  design: 'ribbonTabDesign',
  layout: 'ribbonTabLayout',
  references: 'ribbonTabReferences',
  review: 'ribbonTabReview',
  view: 'ribbonTabView',
  tableDesign: 'ribbonTabTableDesign',
  tableLayout: 'ribbonTabTableLayout',
  pictureFormat: 'ribbonTabPictureFormat',
  shapeFormat: 'ribbonTabShapeFormat',
  headerFooter: 'ribbonTabHeaderFooter',
}

/** Word's picture size limits in twips (0.01"–22") */
export const PICTURE_TWIPS_MIN = 14
export const PICTURE_TWIPS_MAX = 31680
export const clampPictureTwips = (twips: number) =>
  Math.min(PICTURE_TWIPS_MAX, Math.max(PICTURE_TWIPS_MIN, twips))

/** swallows every command when the document is read-only (protected / read mode) */
const NOOP_CHAIN = new Proxy(
  {},
  { get: (_t, prop) => (prop === 'run' ? () => false : () => NOOP_CHAIN) },
) as ChainedCommands

// A+/A- clicks closer together than this coalesce into one trailing apply;
// must sit above burst-click spacing (~100-200ms) yet stay short enough that
// the deferred re-layout still feels attached to the click.
const FONT_STEP_COALESCE_MS = 300

/** Word text highlight colors (OOXML named values) */
const HIGHLIGHTS = [
  'yellow',
  'green',
  'cyan',
  'magenta',
  'blue',
  'red',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'black',
]

const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3]

export { numberPresetLevels } from '../list-presets'

function findNumIdOfKind(blocks: Block[], kind: 'bullet' | 'ordered'): string | null {
  for (const b of blocks) {
    if (b.type === 'listItem' && b.list?.kind === kind) return b.list.numId
  }
  return null
}

// the caret stays in the document, as in Word; a scripted refocus would also
// make Blink cancel any spell-check pass a later selection change interrupts
function keepDocumentFocus(e: ReactMouseEvent<HTMLDivElement>) {
  const target = e.target as HTMLElement
  if (target.closest('input, select, textarea, [contenteditable="true"], .gs-dd')) return
  if (target.closest('button')) e.preventDefault()
}

function RibbonInner({
  actionsRef,
  quickActions,
  trailingActions,
  editor,
  formatState: fs,
  hasDoc,
  blocks,
  allocateNumId,
  createListDef,
  onListDialog,
  documentListPresets,
  recentListPresets,
  onParagraphDialog,
  onPasteDefaults,
  onTableDialog,
  tableGridlines,
  onToggleTableGridlines,
  onFind,
  onReplace,
  styles,
  docDefaults,
  onOpen,
  onSave,
  onSaveAs,
  showAi,
  onToggleAi,
  section,
  onSection,
  activeSection,
  onInsertSectionBreak,
  onPaperSizeAll,
  mirrorMargins,
  onMirrorMargins,
  pageColor,
  onPageColor,
  watermark,
  onWatermark,
  themeFonts,
  onThemeFonts,
  onThemeColors,
  inkTool,
  onInkTool,
  inkPen,
  onInkPen,
  inkHighlighter,
  onInkHighlighter,
  inkCount,
  onInkClearAll,
  onInsertNote,
  sources,
  zoteroNoteFields,
  onAddSource,
  headingPages,
  zoom,
  onZoom,
  onZoomFit,
  onZoomDialog,
  darkPage,
  onDarkPage,
  onAiPreset,
  tabRequest,
  header,
  onHeader,
  onPageNumFormat,
  onInsertField,
  footer,
  onFooter,
  hfEditing,
  onHfAction,
  onHfEdit,
  showMarks,
  onShowMarks,
  showRuler,
  onShowRuler,
  showNav,
  onShowNav,
  showStylesPane,
  onShowStylesPane,
  commentCount,
  openCommentCount,
  resolvedCommentCount,
  onShowComments,
  canComment,
  onNewComment,
  commentAtCaret,
  onDeleteComment,
  onDeleteAllComments,
  onGotoComment,
  trackChanges,
  onTrackChanges,
  spellcheck,
  onSpellcheck,
  revisionDisplay,
  onRevisionDisplay,
  revisionCount,
  onAcceptRevision,
  onRejectRevision,
  onGotoRevision,
  isProtected,
  commentsAllowed,
  trackChangesForced,
  protectActive,
  onProtectDoc,
  onCompare,
  filePath,
  viewMode,
  onViewMode,
  readMode,
  onReadMode,
  showGrid,
  onShowGrid,
  splitView,
  onSplitView,
  onPagePreview,
}: RibbonProps) {
  const { t, lang } = useI18n()
  const collapse = useRibbonCollapse('aidocs.ribbonCollapsed', {
    collapse: t('ribbonCollapse'),
    expand: t('ribbonExpand'),
  })
  // The one-click AI actions need text to work on; grey them out on an empty document
  const docEmpty = !hasDoc || fs.docEmpty
  const [tab, setTab] = useState<RibbonTab>('home')
  const [dropdown, setDropdown] = useState<string | null>(null)
  // null = Automatic: the pen button clears the run colour instead of writing one
  const [penColor, setPenColor] = useState<string | null>('FF0000')
  const [penHighlight, setPenHighlight] = useState('yellow')
  const [painter, setPainter] = useState<PainterState | null>(null)
  /** formatting stored by the painter button or the copy-formatting shortcut, pasted by its twin */
  const formatClipRef = useRef<PainterFormat | null>(null)
  /** painter state before each of the last two button clicks: a double-click's own clicks already toggled it */
  const painterPressRef = useRef<boolean[]>([])
  const fontStepRef = useRef<{
    pending: number | null
    applied: number | null
    timer: number | null
    // editor snapshot the deferred apply validates against (stale-apply guard)
    anchor: number
    head: number
    doc: PMNode | null
  }>({ pending: null, applied: null, timer: null, anchor: -1, head: -1, doc: null })
  /** Font / size combo drafts: null = box mirrors the selection; Esc and blur drop the draft */
  const [fontDraft, setFontDraft] = useState<string | null>(null)
  const [fontHighlight, setFontHighlight] = useState<string | null>(null)
  const [sizeDraft, setSizeDraft] = useState<string | null>(null)
  const fontInputRef = useRef<HTMLInputElement>(null)
  const fontCompletionRef = useRef<[number, number] | null>(null)
  // a press inside a combo wrap (caret, list) blurs the input; that blur must not close the list
  const comboPressRef = useRef(false)
  const lastRegularTab = useRef<(typeof TABS)[number]>('home')
  const wasInTable = useRef(false)
  const wasInImage = useRef(false)
  const wasInHf = useRef(false)
  /** Picture Format → remove background / crop dialogs */
  const [pictureDialog, setPictureDialog] = useState<'cutout' | 'crop' | null>(null)
  // Word's Lock aspect ratio is per picture: a newly selected picture starts locked
  const [lockAspect, setLockAspect] = useState(true)
  useEffect(() => setLockAspect(true), [fs.imageKey])

  useEffect(() => {
    if (!tabRequest) return
    if ((TABS as readonly string[]).includes(tabRequest.tab)) {
      const requested = tabRequest.tab as (typeof TABS)[number]
      lastRegularTab.current = requested
      setTab(requested)
      setDropdown(null)
    } else if (tabRequest.tab === 'tableDesign') {
      setTab('tableDesign')
      setDropdown(null)
    }
  }, [tabRequest])

  // Unified dismissal: a press anywhere outside the open panel closes it (plus
  // window blur / shell chrome presses). The [data-rb-panel] element exists in
  // the DOM only while a dropdown is open, and its parent element is the wrap
  // that also holds the trigger button — so a press on the open dropdown's own
  // trigger counts as "inside" and falls through to the trigger's onClick
  // toggle (closing it) instead of being treated as an outside press.
  useDismissablePopover(dropdown != null, () => setDropdown(null), {
    inside: () =>
      Array.from(document.querySelectorAll('[data-rb-panel]')).flatMap((panel) => [
        panel,
        panel.parentElement,
      ]),
  })

  // leaving the Draw tab always drops back to text editing, so the drawing
  // overlay never swallows clicks while its controls are off-screen
  useEffect(() => {
    if (tab !== 'draw' && inkTool !== 'select') onInkTool('select')
  }, [tab, inkTool, onInkTool])

  // a focused textbox sub-editor receives text/paragraph formatting instead
  // of the main editor (Word: ribbon acts on the shape's text while inside it)
  const sub = fs.sub
  const ed = sub ?? editor
  // read-only (Restrict Editing / Read Mode): every edit command is fenced here,
  // button disabled states are only the visual layer on top
  const canEdit = hasDoc && fs.editable
  const chain = () => (canEdit ? ed.chain().focus() : NOOP_CHAIN)
  const inTable = fs.inTable

  // Word: the caret entering an existing table only shows Table Design /
  // Table Layout; the active tab changes only for a freshly inserted table
  useEffect(() => {
    if (inTable && !wasInTable.current) {
      wasInTable.current = true
    } else if (!inTable && wasInTable.current) {
      wasInTable.current = false
      setDropdown(null)
      setTab((current) =>
        TABLE_TABS.includes(current as (typeof TABLE_TABS)[number])
          ? lastRegularTab.current
          : current,
      )
    }
  }, [inTable])

  // ---- Header & Footer (contextual tab while a strip editor is open, same mechanism) ----
  const inHf = hfEditing !== null
  useEffect(() => {
    if (inHf && !wasInHf.current) {
      wasInHf.current = true
      setDropdown(null)
      setTab('headerFooter')
    } else if (!inHf && wasInHf.current) {
      wasInHf.current = false
      setDropdown(null)
      setTab((current) => (current === 'headerFooter' ? lastRegularTab.current : current))
    }
  }, [inHf])

  // ---- Picture Format (contextual tab when an image block is selected, same mechanism as tables) ----
  const inImage = !sub && fs.imageSelected
  const imageDataUrl = inImage ? fs.imageDataUrl : null

  useEffect(() => {
    if (inImage && !wasInImage.current) {
      wasInImage.current = true
      setDropdown(null)
      setTab('pictureFormat')
    } else if (!inImage && wasInImage.current) {
      wasInImage.current = false
      setDropdown(null)
      setPictureDialog(null)
      setTab((current) => (current === 'pictureFormat' ? lastRegularTab.current : current))
    }
  }, [inImage])

  // ---- Shape Format (contextual tab when a floating box is selected, same mechanism) ----
  // Unlike the picture tab this one survives `sub`: double-clicking into the
  // shape's text keeps the object selected in the main editor, and Word leaves
  // Shape Format standing throughout — dropping it there is what forced a trip
  // back to Home to change so much as the weight of the text just typed.
  const inShape = fs.textboxSelected
  const shapeIsLine = !!fs.shapePrst?.startsWith('line')
  const wasInShape = useRef(false)

  useEffect(() => {
    if (inShape && !wasInShape.current) {
      wasInShape.current = true
      setDropdown(null)
      setTab('shapeFormat')
    } else if (!inShape && wasInShape.current) {
      wasInShape.current = false
      setDropdown(null)
      setTab((current) => (current === 'shapeFormat' ? lastRegularTab.current : current))
    }
  }, [inShape])

  /** apply fill/outline to the selected floating box (first box of the node) */
  const setShapeStyle = (patch: { fill?: string | null; borderColor?: string | null }) => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const boxes = attrs?.textboxes as TextboxDisplay[] | null
    if (!Array.isArray(boxes) || boxes.length === 0) return
    const box = { ...boxes[0] }
    if ('fill' in patch) {
      if (patch.fill) box.fill = patch.fill
      else delete box.fill
    }
    if ('borderColor' in patch) {
      if (patch.borderColor) box.borderColor = patch.borderColor
      else delete box.borderColor
    }
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { textboxes: [box, ...boxes.slice(1)] })
      .run()
  }

  /**
   * Rewrite every run and paragraph of the selected shape. Only for object mode:
   * with the shape selected there is no text selection for a mark command to act
   * on, so Word reformats the whole shape and `editRun`/`editPara` see each part
   * of it in turn. While a sub-editor holds focus the ordinary mark and alignment
   * commands already target the text, and shapeTextCommand routes there instead.
   *
   * Confined to the first box like setShapeStyle: a paragraph anchoring several
   * shapes packs them into one docProtected node, and the ribbon reads and writes
   * only that one — reformatting the rest would hit shapes it is not showing.
   *
   * The node view re-feeds its sub-editors from the new attrs, so the shape
   * repaints even while its text is being edited.
   */
  const setShapeText = (
    editRun?: (run: Run) => Run,
    editPara?: (para: TextboxParaDisplay) => TextboxParaDisplay,
  ) => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const boxes = attrs?.textboxes as TextboxDisplay[] | null
    if (!Array.isArray(boxes) || boxes.length === 0 || boxes[0].readOnly) return
    const box = {
      ...boxes[0],
      paras: boxes[0].paras.map((para) => {
        const withRuns = editRun ? { ...para, runs: para.runs.map(editRun) } : para
        return editPara ? editPara(withRuns) : withRuns
      }),
    }
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { textboxes: [box, ...boxes.slice(1)] })
      .run()
  }

  /** drop a run property rather than storing an explicit "off" Word would have to write out */
  const withRunFlag = (run: Run, key: 'bold' | 'italic' | 'underline', on: boolean): Run => {
    const next = { ...run }
    if (on) next[key] = true
    else delete next[key]
    return next
  }

  /** Shape Format text buttons: the sub-editor when inside the text, the whole shape otherwise */
  const shapeTextCommand = {
    toggleMark: (name: 'bold' | 'italic' | 'underline', active: boolean) => {
      if (sub) chain().toggleMark(name).run()
      else setShapeText((run) => withRunFlag(run, name, !active))
    },
    setColor: (hex: string | null) => {
      if (sub) setTextStyle({ color: hex })
      else
        setShapeText((run) => {
          const next = { ...run }
          if (hex) next.color = hex
          else delete next.color
          return next
        })
    },
    setAlign: (align: 'left' | 'center' | 'right' | 'justify') => {
      if (sub) setSelectionAlign(ed, align)
      else setShapeText(undefined, (para) => ({ ...para, align }))
    },
  }

  const shapeTextActive = {
    bold: sub ? fs.bold : fs.shapeTextBold,
    italic: sub ? fs.italic : fs.shapeTextItalic,
    underline: sub ? fs.underline : fs.shapeTextUnderline,
    color: sub ? fs.textColor : fs.shapeTextColor,
    align: sub ? fs.align : fs.shapeTextAlign,
  }

  /**
   * Replace the selected image's bytes (shared by Replace Picture / remove background / crop).
   * Original images (docxIndex set) swap bytes in place via the imageReplace patch: the
   * drawing XML survives, so wrap/position/docxIndex — and with them the Position gallery —
   * keep working. Images not yet saved (genImage) just update their pending payload.
   * Display size keeps the current width; height adapts to the new image's aspect ratio.
   */
  const applyPictureBytes = async (dataUrl: string) => {
    if (!canEdit) return
    const m = /^data:(image\/(?:png|jpeg|gif));base64,(.*)$/s.exec(dataUrl)
    if (!m) return
    const attrs = editor.getAttributes('docProtected')
    if (attrs?.blockType !== 'image') return
    try {
      const natural = await imageSizeOf(dataUrl)
      const currentW = Number(attrs.imageWidthPx) || Math.min(natural.width, 620)
      const w = Math.max(1, Math.round(currentW))
      const h = Math.max(1, Math.round((currentW * natural.height) / natural.width))
      const isOriginal = attrs.docxIndex !== null && attrs.docxIndex !== undefined
      editor
        .chain()
        .focus()
        .updateAttributes('docProtected', {
          imageDataUrl: dataUrl,
          imageWidthPx: w,
          imageHeightPx: h,
          // The new bytes are the full picture (crop/cutout bake destructively) and
          // the replace pipeline strips a:srcRect on save — drop a Word-authored
          // crop/fill window or it would keep clipping the new image until reload
          imageCrop: null,
          imageFillRect: null,
          ...(isOriginal
            ? { imageReplace: { base64: m[2], mime: m[1] } }
            : { genImage: { base64: m[2], mime: m[1], widthPx: w, heightPx: h } }),
        })
        .run()
    } catch {
      /* image decode failed: keep the original untouched */
    }
  }

  const replacePicture = async () => {
    const picked = await window.desktop.pickImage()
    if (!picked) return
    await applyPictureBytes(`data:${picked.mime};base64,${picked.base64}`)
  }

  const rotatePicture = (deltaDeg: number) => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    if (attrs?.blockType !== 'image') return
    const next = ((((Number(attrs.imageRotDeg) || 0) + deltaDeg) % 360) + 360) % 360
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageRotDeg: next || null })
      .run()
  }

  const flipPicture = (axis: 'h' | 'v') => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    if (attrs?.blockType !== 'image') return
    const key = axis === 'h' ? 'imageFlipH' : 'imageFlipV'
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { [key]: !attrs[key] })
      .run()
  }

  /** Set one side of the picture; with the aspect ratio locked the other side follows */
  const setPictureSize = (dim: 'w' | 'h', twips: number) => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const w = Number(attrs?.imageWidthPx)
    const h = Number(attrs?.imageHeightPx)
    if (attrs?.blockType !== 'image' || !w || !h || !(twips > 0)) return
    const px = twipsToPx(clampPictureTwips(twips))
    const side = Math.max(1, Math.round(px))
    const next =
      dim === 'w'
        ? {
            imageWidthPx: side,
            imageHeightPx: lockAspect ? Math.max(1, Math.round((px * h) / w)) : h,
          }
        : {
            imageWidthPx: lockAspect ? Math.max(1, Math.round((px * w) / h)) : w,
            imageHeightPx: side,
          }
    editor.chain().focus().updateAttributes('docProtected', next).run()
  }

  /** Word's Reset Size: the picture's own pixel size at 96 dpi. A picture wider
   *  than the text column is fitted to it (the page cannot show the overflow). */
  const resetPictureSize = async () => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const url = attrs?.imageDataUrl as string | null
    if (attrs?.blockType !== 'image' || !url) return
    try {
      const natural = await imageSizeOf(url)
      const scale = Math.min(1, sectionContentWidthPx / natural.width)
      editor
        .chain()
        .focus()
        .updateAttributes('docProtected', {
          imageWidthPx: Math.max(1, Math.round(natural.width * scale)),
          imageHeightPx: Math.max(1, Math.round(natural.height * scale)),
        })
        .run()
    } catch {
      /* decode failed: leave as is */
    }
  }

  const runTableCommand = (command: Command) => {
    if (!canEdit) return
    editor.view.focus()
    enterSelectedTable(editor.state, editor.view.dispatch)
    command(editor.state, editor.view.dispatch)
  }

  // ---- Table borders / vertical alignment / row height & column width ----
  const [borderColor, setBorderColor] = useState('000000')
  const [borderSz, setBorderSz] = useState(4) // 1/8 pt:4 = 0.5pt
  const [borderStyle, setBorderStyle] = useState('single')
  const sectionContentWidthPx = section
    ? Math.max(1, (section.pageWidth - section.marginLeft - section.marginRight) / 15)
    : 624
  const maxRowHeightTwips = section
    ? Math.max(1, section.pageHeight - section.marginTop - section.marginBottom)
    : 13200

  const [lastBorderMode, setLastBorderMode] = useState<TableBorderMode>('bottom')
  const LastBorderIcon = TABLE_BORDER_ITEMS[lastBorderMode].Icon
  const applyCellBorders = (mode: TableBorderMode) => {
    setLastBorderMode(mode)
    runTableCommand(
      setSelectionBorders(mode, { style: borderStyle, szEighths: borderSz, color: borderColor }),
    )
  }

  /** Set row height for selected rows (twips; null = auto) */
  const applyRowHeight = (height: number | null) => {
    if (!canEdit) return
    editor.view.focus()
    if (!enterSelectedTable(editor.state, editor.view.dispatch)) return
    const { state, view } = editor
    const rect = selectedRect(state)
    const twips = height && height > 0 ? Math.min(height, maxRowHeightTwips) : null
    let tr = state.tr
    rect.table.forEach((rowNode, offset, idx) => {
      if (idx < rect.top || idx >= rect.bottom) return
      tr = tr.setNodeMarkup(rect.tableStart + offset, undefined, {
        ...rowNode.attrs,
        heightTwips: twips,
      })
    })
    view.dispatch(tr)
  }

  /** Set column width for selected columns (twips): writes the matching colwidth slot of every cell in the column */
  const applyColumnWidth = (twips: number | null) => {
    if (!canEdit || !twips || twips <= 0) return
    editor.view.focus()
    if (!enterSelectedTable(editor.state, editor.view.dispatch)) return
    const px = Math.max(1, Math.round(twipsToPx(twips)))
    setSelectedColumnWidth(px, sectionContentWidthPx)(editor.state, editor.view.dispatch)
  }

  /** Current cell properties (echoed in the size inputs) */
  const activeCellInfo =
    fs.cellKey === null
      ? null
      : {
          key: fs.cellKey,
          heightTwips: fs.cellHeightTwips,
          widthTwips: fs.cellWidthPx === null ? null : pxToTwips(fs.cellWidthPx),
          vAlign: fs.cellVAlign,
        }

  const tableAttrs = inTable ? editor.getAttributes('docTable') : {}
  const tableHeader = (() => {
    if (!inTable) return { enabled: false, active: false }
    const cells = tableCellsSelection(editor.state)
    return repeatHeaderState(
      cells ? editor.state.apply(editor.state.tr.setSelection(cells)) : editor.state,
    )
  })()
  const tableLook = (tableAttrs.tblLook as TableLook | null) ?? {
    firstRow: true,
    lastRow: false,
    firstColumn: true,
    lastColumn: false,
    bandedRows: true,
    bandedColumns: false,
  }
  const tableAutoFitMode: TableAutoFitMode =
    tableAttrs.tblAutoFit === 'contents' || tableAttrs.tblAutoFit === 'window'
      ? tableAttrs.tblAutoFit
      : 'fixed'
  const tableAutoFitLabel =
    TABLE_AUTO_FIT_OPTIONS.find(([mode]) => mode === tableAutoFitMode)?.[1] ??
    'ribbonFixedColumnWidth'

  const activeCharStyleId = fs.charStyleId
  const galleryStyles = styles ?? FALLBACK_STYLES
  // the used set is folded to a string so an unchanged set keeps the memoized entries
  const usedKey = useMemo(
    () => [...collectUsedStyleIds(fs.doc, galleryStyles)].sort().join('\u0001'),
    [fs.doc, galleryStyles],
  )
  const galleryEntries = useMemo(
    () => quickStyleEntries(galleryStyles, new Set(usedKey ? usedKey.split('\u0001') : [])),
    [galleryStyles, usedKey],
  )
  const activeStyleKey = activeStyleKeyOf(fs, styles)

  // Style gallery overflow: cards that don't fit wrap onto a second row that
  // the fixed-height gallery clips (whole cards only, never a half-cut one),
  // and a "more styles" expander appears whenever cards are hidden. The
  // gallery is then capped right after the last visible card so the expander
  // hugs it instead of floating at the group's far edge.
  const styleGalleryRef = useRef<HTMLDivElement | null>(null)
  const [styleGalleryOverflow, setStyleGalleryOverflow] = useState(false)
  useLayoutEffect(() => {
    const el = styleGalleryRef.current
    if (!el || typeof ResizeObserver === 'undefined') return
    const check = () => {
      // measure the natural (uncapped) layout at the current wrapper width
      el.style.maxWidth = ''
      const cards = Array.from(el.children) as HTMLElement[]
      const firstRow = cards.filter((c) => c.offsetTop === cards[0]?.offsetTop)
      const overflow = firstRow.length < cards.length
      if (overflow) {
        const last = firstRow[firstRow.length - 1]
        el.style.maxWidth = `${last.offsetLeft - firstRow[0].offsetLeft + last.offsetWidth}px`
      }
      setStyleGalleryOverflow(overflow)
    }
    check()
    const ro = new ResizeObserver(check)
    // observe the wrapper, not the gallery: once capped, the gallery no longer
    // resizes with the window, so it would never re-trigger the observer
    ro.observe(el.parentElement ?? el)
    return () => ro.disconnect()
    // re-check when the card set can change, and after the expander mounts or
    // unmounts (it takes row width, which can change how many cards fit)
  }, [tab, galleryEntries.length, lang, styleGalleryOverflow])

  const currentSize = fs.fontSizePt
  // computed unconditionally (not inside the dropdown render): cheap, and the
  // render-isolation test uses fontFamiliesFor calls as its render probe
  const fontFamilies = fontFamiliesFor(lang)
  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // unset align follows the paragraph direction: start is left in LTR, right in RTL
  const activeAlign = fs.align ?? (fs.bidi ? 'right' : 'left')
  // like Word, direction buttons appear only with an RTL UI language or RTL paragraphs at hand
  const showDirection = isRtlUiLang(lang) || fs.selectionBidi
  const activeSpacing = fs.lineSpacing

  /** merge new attrs into the docTextStyle mark, preserving the rest.
   * Only the patch is passed: setMark merges per existing mark and with the caret's
   * stored mark. Rebuilding from getAttributes read-back dropped the previous call's
   * value on a collapsed cursor (stored-mark changes don't re-render). */
  const setTextStyle = (patch: Record<string, unknown>) => {
    chain().setMark('docTextStyle', patch).run()
    setDropdown(null)
  }

  const currentFont = fs.fontFamily
  // Word's font box: an East Asian face fills every rFonts slot, a Latin face only
  // w:ascii/w:hAnsi so the CJK font survives
  const setFont = (name: string) => {
    setTextStyle(
      isEastAsianFontName(name)
        ? { font: name, fontAscii: name, eastAsiaFont: name, eaSlotEmpty: false }
        : { fontAscii: name },
    )
  }
  const fontMenuItems: Array<{ key: string; name: string; section: 'b' | 's' }> = [
    ...fontFamilies.map((name) => ({ key: `b:${name}`, name, section: 'b' as const })),
    ...systemFontFamilies.map((name) => ({ key: `s:${name}`, name, section: 's' as const })),
  ]
  const openFontMenu = () => {
    if (dropdown === 'fontFamily') return
    loadSystemFonts()
    setDropdown('fontFamily')
  }
  const resetFontBox = () => {
    setFontDraft(null)
    setFontHighlight(null)
  }
  const onFontTyped = (e: ChangeEvent<HTMLInputElement>) => {
    const input = e.currentTarget
    const typed = input.value
    const native = e.nativeEvent as InputEvent
    const q = typed.trim().toLowerCase()
    const match = q ? fontMenuItems.find((it) => it.name.toLowerCase().startsWith(q)) : undefined
    // inline completion only on plain insertions, so Backspace shortens instead of re-completing
    if (
      match &&
      native.inputType === 'insertText' &&
      !native.isComposing &&
      match.name.length > typed.length &&
      input.selectionStart === typed.length
    ) {
      input.value = match.name
      input.setSelectionRange(typed.length, match.name.length)
      fontCompletionRef.current = [typed.length, match.name.length]
      setFontDraft(match.name)
    } else setFontDraft(typed)
    setFontHighlight(match?.key ?? null)
    openFontMenu()
  }
  const onFontKeyDown = (e: ReactKeyboardEvent<HTMLInputElement>) => {
    if (e.nativeEvent.isComposing) return
    if (e.key === 'Enter') {
      e.preventDefault()
      const picked =
        dropdown === 'fontFamily' && fontHighlight
          ? fontMenuItems.find((it) => it.key === fontHighlight)?.name
          : undefined
      const name = (picked ?? e.currentTarget.value).trim()
      resetFontBox()
      // an unchanged name is a no-op: re-applying the East Asian face the box
      // shows on CJK text would overwrite the run's Latin font
      if (name && name !== currentFont) setFont(name)
      else {
        setDropdown(null)
        chain().run()
      }
    } else if (e.key === 'Escape') {
      e.preventDefault()
      e.stopPropagation()
      resetFontBox()
      setDropdown(null)
      chain().run()
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault()
      if (fontMenuItems.length === 0) return
      openFontMenu()
      const at = fontMenuItems.findIndex((it) => it.key === fontHighlight)
      const base = at === -1 ? fontMenuItems.findIndex((it) => it.name === currentFont) : at
      const next =
        e.key === 'ArrowDown' ? Math.min(base + 1, fontMenuItems.length - 1) : Math.max(base - 1, 0)
      setFontHighlight(fontMenuItems[next].key)
      setFontDraft(fontMenuItems[next].name)
    }
  }
  useLayoutEffect(() => {
    const range = fontCompletionRef.current
    if (range && fontInputRef.current) fontInputRef.current.setSelectionRange(range[0], range[1])
    fontCompletionRef.current = null
  })
  useEffect(() => {
    if (!fontHighlight || dropdown !== 'fontFamily') return
    for (const el of document.querySelectorAll<HTMLElement>('[data-font-key]'))
      if (el.dataset.fontKey === fontHighlight) {
        el.scrollIntoView?.({ block: 'nearest' })
        break
      }
  }, [fontHighlight, dropdown])
  const markComboPress = () => {
    comboPressRef.current = true
    window.setTimeout(() => (comboPressRef.current = false), 0)
  }
  /** blur without Enter: Word drops the draft and keeps the selection's value */
  const onComboBlur = (reset: () => void, menu: string) => {
    setInactiveSelectionShown(ed, false)
    reset()
    if (!comboPressRef.current) setDropdown((v) => (v === menu ? null : v))
  }
  const commitSize = (text: string) => {
    const pt = parseFontSize(text, lang)
    setSizeDraft(null)
    // Enter applies even an unchanged value: it normalizes a mixed-size selection
    if (pt !== null) setTextStyle({ sizeHalfPoints: Math.round(pt * 2) })
    else {
      setDropdown(null)
      chain().run()
    }
  }

  /** apply paragraph-level attrs to every paragraph in the selection (textbox sub-editor included) */
  const setParaAttr = (attrs: Record<string, unknown>) => {
    if (canEdit) setParaAttrs(ed, attrs)
    setDropdown(null)
  }

  const applyStyle = (info: StyleInfo) => {
    if (!canEdit) return
    applyGalleryStyle(editor, sub, info, styles, activeCharStyleId)
  }

  /** Style cards, shared by the inline gallery and its overflow menu */
  const renderStyleCards = (inMenu: boolean) => {
    const apply = (info: StyleInfo) => {
      applyStyle(info)
      if (inMenu) setDropdown(null)
    }
    return (
      <>
        {galleryEntries.map((info) => {
          const isChar = info.type === 'character'
          const key = styleKeyOf(info)
          const label = styleLabel(info, t)
          return (
            <button
              key={key}
              className={`style-card${isChar ? ' style-card-char' : ''}${activeStyleKey === key ? ' active' : ''}`}
              disabled={!canEdit || (!isChar && !!sub)}
              data-tip={label}
              data-style-id={info.styleId}
              onClick={() => apply(info)}
            >
              <span className="style-card-preview" style={stylePreviewCss(info, docDefaults)}>
                {isChar ? 'Aa' : t('ribbonStylePreview')}
              </span>
              <span className="style-card-label">{label}</span>
            </button>
          )
        })}
      </>
    )
  }

  const toggleList = (kind: 'bullet' | 'ordered') => {
    if (sub) return // textboxes have no list numbering
    if (editor.isActive('docListItem', { kind })) {
      chain().setNode('docParagraph').run()
      return
    }
    // reuse the numId of an existing same-kind instance in the body; otherwise adopt a document definition / create one (writes numbering.xml)
    const numId = findNumIdOfKind(blocks, kind) ?? allocateNumId?.(kind) ?? null
    chain().setNode('docListItem', { kind, numId, ilvl: 0 }).run()
  }

  /** the gallery "None" card: drop list formatting, back to a plain paragraph */
  const clearList = () => {
    if (sub) return
    if (editor.isActive('docListItem')) chain().setNode('docParagraph').run()
  }

  /** Custom levels picked in the gallery/dialog → create a definition and apply it to the current paragraph */
  const applyListPreset = (levels: CustomNumberingLevel[], recentKind?: ListPresetKind) => {
    if (sub) return
    const numId = createListDef?.(levels) ?? null
    if (!numId) return
    if (recentKind) rememberListPreset(recentKind, levels)
    const kind = levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered'
    const ilvl = editor.isActive('docListItem')
      ? Number(editor.getAttributes('docListItem').ilvl) || 0
      : 0
    chain().setNode('docListItem', { kind, numId, ilvl }).run()
  }

  const changeListLevel = (ilvl: number) => {
    if (sub || !editor.isActive('docListItem')) return
    chain().updateAttributes('docListItem', { ilvl }).run()
    setDropdown(null)
  }

  const toggleParaSpace = (side: 'spaceBefore' | 'spaceAfter') => {
    if (canEdit)
      toggleParaSpaceImpl(ed, side, side === 'spaceBefore' ? fs.spaceBefore : fs.spaceAfter)
    setDropdown(null)
  }

  const listLevelRow = (
    <div className="list-gallery-levels">
      <span className="list-gallery-title">{t('ribbonChangeListLevel')}</span>
      {Array.from({ length: 9 }, (_, i) => (
        <button
          key={i}
          className={`list-gallery-level${fs.listBullet || fs.listOrdered ? '' : ' disabled'}${
            editor.isActive('docListItem') &&
            (Number(editor.getAttributes('docListItem').ilvl) || 0) === i
              ? ' selected'
              : ''
          }`}
          disabled={!(fs.listBullet || fs.listOrdered)}
          aria-label={t('appParaOutlineLevelN', { n: String(i + 1) })}
          onClick={() => changeListLevel(i)}
        >
          {i + 1}
        </button>
      ))}
    </div>
  )

  const presetSection = (
    titleKey: StringKey,
    presets: CustomNumberingLevel[][] | undefined,
    kind: ListPresetKind,
    render: (levels: CustomNumberingLevel[]) => ReactNode,
  ) =>
    presets && presets.length > 0 ? (
      <>
        <div className="list-gallery-title">{t(titleKey)}</div>
        {presets.map((levels, i) => (
          <button
            key={`${kind}-${i}`}
            className={`list-gallery-card ${kind === 'multi' ? 'list-gallery-card-multi' : kind === 'bullets' ? 'list-gallery-glyph' : 'list-gallery-preview'}`}
            onClick={() => {
              applyListPreset(levels, kind)
              setDropdown(null)
            }}
          >
            {render(levels)}
          </button>
        ))}
      </>
    ) : null

  const numberPreview = (levels: CustomNumberingLevel[]) =>
    [0, 1, 2].map((k) => (
      <span key={k} className="list-gallery-preview-row">
        <span className="list-gallery-preview-prefix">
          {levels[0].lvlText.replace(
            /%1/g,
            formatNumber((levels[0].start ?? 1) + k, levels[0].numFmt),
          )}
        </span>
        <span className="list-gallery-preview-line" />
      </span>
    ))

  const bulletPreview = (levels: CustomNumberingLevel[]) => (
    <span
      style={{
        fontFamily: levels[0].font,
        color: levels[0].color ? `#${levels[0].color}` : undefined,
      }}
    >
      {levels[0].lvlText}
    </span>
  )

  const multiPreview = (levels: CustomNumberingLevel[]) =>
    [0, 1, 2].map((lvl) => (
      <span key={lvl} style={{ paddingLeft: lvl * 10 }}>
        {previewLevelText(levels, lvl)} ———
      </span>
    ))

  const changeIndent = (delta: 1 | -1) => {
    if (sub || !canEdit) return
    stepParagraphIndent(editor, delta)
  }

  const applyFontStep = (step: (base: number) => number) => {
    // Every applied size change re-paginates the whole document synchronously —
    // ~700ms per click on table-heavy documents — so clicking A+/A- in a burst
    // froze the UI for seconds. Apply the first click immediately (a single
    // click keeps instant feedback); clicks landing inside the coalesce window
    // only advance the pending size, and one trailing apply lays out the final
    // size. `pending` also covers fs.fontSizePt lagging the last apply within
    // the window.
    const st = fontStepRef.current
    const next = step(st.pending ?? currentSize)
    st.pending = next
    if (st.timer === null) {
      st.applied = next
      setTextStyle({ sizeHalfPoints: Math.round(next * 2) })
    } else {
      window.clearTimeout(st.timer)
    }
    // Snapshot after the (possible) leading apply: the deferred apply is only
    // valid while nothing else has touched the editor. A selection move, an
    // undo, or a size set another way each shows up as a selection or document
    // change and must invalidate the pending step instead of being overwritten.
    const target = ed
    st.anchor = target.state.selection.anchor
    st.head = target.state.selection.head
    st.doc = target.state.doc
    st.timer = window.setTimeout(() => {
      st.timer = null
      const pending = st.pending
      st.pending = null
      if (pending === null || pending === st.applied || !canEdit) return
      if (
        target.state.selection.anchor !== st.anchor ||
        target.state.selection.head !== st.head ||
        target.state.doc !== st.doc
      )
        return
      st.applied = pending
      // deliberately no focus(): a deferred apply must never pull focus back
      target
        .chain()
        .setMark('docTextStyle', { sizeHalfPoints: Math.round(pending * 2) })
        .run()
    }, FONT_STEP_COALESCE_MS)
  }

  /** A+/A- and ⇧⌘. / ⇧⌘,: walk the UI language's size list, tens above it, points below */
  const stepFontSize = (dir: 1 | -1) => applyFontStep((base) => stepSizeInList(base, dir, lang))

  /** Word's ⌘] / ⌘[: exactly one point, within Word's 1–1638pt range */
  const nudgeFontSize = (dir: 1 | -1) =>
    applyFontStep((base) => Math.min(Math.max(Math.round(base) + dir, 1), 1638))

  useEffect(() => {
    if (!actionsRef) return
    actionsRef.current.stepFontSize = stepFontSize
    actionsRef.current.nudgeFontSize = nudgeFontSize
  })

  const toggleVertAlign = (kind: 'superscript' | 'subscript') => {
    setTextStyle({ vertAlign: fs.vertAlign === kind ? null : kind })
  }

  /** format painter pickup: the formatting at the caret / of the selection's first run */
  const pickUpFormat = (): PainterFormat | null => {
    if (!canEdit) return null
    const { state } = editor
    const { $from, $head, from, to, empty } = state.selection
    // Word picks up the FIRST character's formatting of a range selection (a
    // triple-clicked paragraph whose last run is plain must still pick up the
    // leading run's look); a collapsed caret reads the marks at the caret.
    const picked: Mark[] = []
    if (empty) {
      picked.push(...$head.marks())
    } else {
      let found = false
      state.doc.nodesBetween(from, to, (node) => {
        if (found) return false
        if (node.isText) {
          found = true
          picked.push(...node.marks)
          return false
        }
        return true
      })
    }
    // Paragraph formatting is picked up per Word's ¶-mark rule: a caret pickup
    // or a cross-paragraph selection carries the block identity (heading
    // level / list numbering / styleId) — which then applies to whole target
    // paragraphs. A PARTIAL in-paragraph drag copies character formatting
    // only — but a selection covering the paragraph's ENTIRE content counts
    // as including the ¶ mark, exactly like Word's triple-click ("select whole
    // paragraph → painter" dropped line spacing/indents while a caret pickup
    // carried them — backwards to any user).
    const { $to } = state.selection
    const coversWholeParagraph =
      !empty &&
      $from.parent.isTextblock && // AllSelection's parent is the doc
      $from.sameParent($to) &&
      $from.parentOffset === 0 &&
      $to.parentOffset === $to.parent.content.size
    const includesParaMark = empty || !$from.sameParent($to) || coversWholeParagraph
    const marks = picked
      .filter((m) => PAINTER_MARK_TYPES.includes(m.type.name) && m.type.name !== 'docTextStyle')
      .map((m) => ({ type: m.type.name, attrs: { ...m.attrs } as Record<string, unknown> }))
    const tsMark = picked.find((m) => m.type.name === 'docTextStyle')
    const ts: Record<string, unknown> = { ...(tsMark?.attrs ?? {}) }
    // raw rPr pass-through belongs to the source run; stamping it on foreign
    // runs would smuggle unmodeled properties across the document
    delete ts.rawRPr
    if (!includesParaMark) {
      // Char-only brush: resolve the EFFECTIVE character formatting (direct
      // marks → character style → paragraph style → docDefaults) and record it
      // as direct formatting, so the brush reproduces what the source LOOKS
      // like even when that look comes from a style. Without this, picking up
      // plain body text (no marks at all) and brushing heading-styled text
      // changes nothing. When the block travels (¶ pickup) it carries the
      // style itself, so no resolved values are stamped there.
      const styleDisplayOf = (id: unknown) =>
        typeof id === 'string' && id ? styles?.get(id)?.display : undefined
      const charStyle = styleDisplayOf(tsMark?.attrs.styleId)
      const paraStyle = styleDisplayOf($from.parent.attrs.styleId)
      for (const t of ['bold', 'italic', 'underline', 'strike'] as const) {
        const styleFlag =
          charStyle?.[t] ??
          paraStyle?.[t] ??
          (t === 'bold' ? docDefaults?.bold : t === 'italic' ? docDefaults?.italic : undefined)
        if (styleFlag && !picked.some((m) => m.type.name === t)) marks.push({ type: t, attrs: {} })
      }
      ts.sizeHalfPoints ??=
        charStyle?.sizeHalfPoints ??
        paraStyle?.sizeHalfPoints ??
        docDefaults?.sizeHalfPoints ??
        null
      ts.color ??= charStyle?.color ?? paraStyle?.color ?? docDefaults?.color ?? null
      ts.fontAscii ??=
        charStyle?.fontAscii ?? paraStyle?.fontAscii ?? docDefaults?.asciiFont ?? null
      if (ts.font == null) {
        // an empty-EA-theme-slot backfill face is not a document font choice — don't stamp it
        if (charStyle?.font && !charStyle.eaSlotEmpty) ts.font = charStyle.font
        else if (paraStyle?.font && !paraStyle.eaSlotEmpty) ts.font = paraStyle.font
        else if (docDefaults?.eastAsiaFont && !docDefaults.eaSlotEmpty)
          ts.font = docDefaults.eastAsiaFont
      }
      ts.csFont ??= charStyle?.csFont ?? paraStyle?.csFont ?? null
      ts.charSpacingTwips ??= charStyle?.charSpacingTwips ?? paraStyle?.charSpacingTwips ?? null
    }
    if (Object.values(ts).some((v) => v != null)) marks.push({ type: 'docTextStyle', attrs: ts })
    const para = $from.parent
    let block: PainterFormat['block'] = null
    const extra = PAINTER_BLOCK_EXTRA[para.type.name]
    if (includesParaMark && extra) {
      const attrs: Record<string, unknown> = {}
      for (const k of [...PAINTER_PARA_KEYS, ...extra]) attrs[k] = para.attrs[k]
      block = { type: para.type.name, attrs }
    }
    return { marks, block }
  }
  const pickUpRef = useRef(pickUpFormat)
  pickUpRef.current = pickUpFormat

  /** arm the painter: a single click brushes once, a double-click (locked) keeps brushing until Esc */
  const armPainter = (locked: boolean) => {
    const fmt = pickUpFormat()
    if (!fmt) return
    formatClipRef.current = fmt
    setPainter({ ...fmt, locked })
  }
  const onPainterClick = () => {
    painterPressRef.current = [...painterPressRef.current.slice(-1), !!painter]
    if (painter) setPainter(null)
    else armPainter(false)
  }
  const onPainterDoubleClick = () => {
    const wasOn = painterPressRef.current[0] ?? !!painter
    painterPressRef.current = []
    if (wasOn) setPainter(null)
    else armPainter(true)
  }

  // Word's copy / paste formatting chords (⇧⌘C / ⇧⌘V; off the Mac the paste
  // chord carries Alt because Ctrl+Shift+V is the paste-and-match-style menu
  // accelerator). Paste targets the selection, or the word at the caret.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey) || !e.shiftKey) return
      const copy = e.code === 'KeyC' && !e.altKey
      const paste = e.code === 'KeyV' && e.altKey === !IS_MAC
      if (!copy && !paste) return
      const el = document.activeElement
      if (el && el !== document.body && !el.closest('.ProseMirror')) return
      if (!editor.isEditable || getActiveSubEditor()) return
      e.preventDefault()
      if (copy) {
        formatClipRef.current = pickUpRef.current()
        return
      }
      const fmt = formatClipRef.current
      if (!fmt) return
      const { from, to, empty, $from } = editor.state.selection
      if (!empty) {
        applyPainterFormat(editor, fmt, from, to, null)
        return
      }
      const word = painterWordRangeAt($from)
      if (word) applyPainterFormat(editor, fmt, word.from, word.to, from)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editor])

  useEffect(() => {
    if (!painter) return
    let selectingWithMouse = false
    let downAt: { x: number; y: number } | null = null
    let done = false
    let keyboardTimer: ReturnType<typeof setTimeout> | null = null
    // The selection the painter must not brush: the pickup selection when
    // armed, then whatever the previous stroke left behind (locked mode). Only
    // a selection that has since MOVED is a target gesture (without this, any
    // stray selectionUpdate right after arming brushes the source itself)
    let idle = { from: editor.state.selection.from, to: editor.state.selection.to }

    const applyRange = (from: number, to: number, caretAfter: number | null) => {
      if (done || !editor.isEditable) return
      done = true
      if (keyboardTimer) clearTimeout(keyboardTimer)
      if (!painter.locked) setPainter(null)
      applyPainterFormat(editor, painter, from, to, caretAfter)
      if (painter.locked) {
        idle = { from: editor.state.selection.from, to: editor.state.selection.to }
        done = false
      }
    }

    const onMouseDown = (event: MouseEvent) => {
      if (!editor.view.dom.contains(event.target as globalThis.Node)) return
      selectingWithMouse = true
      downAt = { x: event.clientX, y: event.clientY }
      if (keyboardTimer) clearTimeout(keyboardTimer)
    }
    const onMouseUp = (event: MouseEvent) => {
      if (!selectingWithMouse || !downAt) return
      selectingWithMouse = false
      const press = downAt
      downAt = null
      const dist = Math.abs(event.clientX - press.x) + Math.abs(event.clientY - press.y)
      requestAnimationFrame(() => {
        if (done || !editor.isEditable) return
        const { from, to } = editor.state.selection
        const moved = from !== idle.from || to !== idle.to
        if (from !== to && moved) {
          applyRange(from, to, null)
          return
        }
        if (dist >= 5) return
        // A plain click brushes the clicked word. The position comes from the
        // press coordinates, not the selection: a fast click into a blurred
        // editor can reach this frame before ProseMirror has placed the caret,
        // and reading the stale selection here used to brush the source itself.
        const hit = editor.view.posAtCoords({ left: press.x, top: press.y })
        if (!hit) return
        const word = painterWordRangeAt(editor.state.doc.resolve(hit.pos))
        if (word) applyRange(word.from, word.to, hit.pos)
      })
    }
    const onSelectionUpdate = () => {
      if (selectingWithMouse || done) return
      if (keyboardTimer) clearTimeout(keyboardTimer)
      keyboardTimer = setTimeout(() => {
        const { from, to } = editor.state.selection
        if (from === idle.from && to === idle.to) return
        if (from !== to) applyRange(from, to, null)
      }, 180)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setPainter(null)
    }

    // Word-style paintbrush cursor over the text area while the painter is armed
    editor.view.dom.classList.add('doc-painter-cursor')
    editor.view.dom.addEventListener('mousedown', onMouseDown)
    window.addEventListener('mouseup', onMouseUp)
    window.addEventListener('keydown', onKeyDown, true)
    editor.on('selectionUpdate', onSelectionUpdate)
    return () => {
      if (keyboardTimer) clearTimeout(keyboardTimer)
      editor.view.dom.classList.remove('doc-painter-cursor')
      editor.view.dom.removeEventListener('mousedown', onMouseDown)
      window.removeEventListener('mouseup', onMouseUp)
      window.removeEventListener('keydown', onKeyDown, true)
      editor.off('selectionUpdate', onSelectionUpdate)
    }
  }, [painter, editor])

  const changeCase = (mode: CaseMode) => {
    if (!canEdit) return
    applyCase(ed, mode)
    setDropdown(null)
  }

  const [pasteSpecial, setPasteSpecial] = useState(false)
  const pasteAs = (mode?: PasteMode) => {
    setDropdown(null)
    if (canEdit) void pasteFromClipboard(ed, mode)
  }

  const clipboard = (action: 'cut' | 'copy') => {
    if (action === 'cut' && !canEdit) return
    document.execCommand(action)
    ed.commands.focus()
  }

  const markBtn = (name: string, active: boolean, title: string, label: ReactNode) => (
    <button
      className={`rb-icon ${active ? 'active' : ''}`}
      disabled={!canEdit}
      data-tip={title}
      aria-label={title}
      onClick={() => chain().toggleMark(name).run()}
    >
      {label}
    </button>
  )

  const shapeMarkBtn = (
    name: 'bold' | 'italic' | 'underline',
    active: boolean,
    title: string,
    label: ReactNode,
  ) => (
    <button
      className={`rb-icon ${active ? 'active' : ''}`}
      disabled={!canEdit}
      data-tip={title}
      aria-label={title}
      onClick={() => shapeTextCommand.toggleMark(name, active)}
    >
      {label}
    </button>
  )

  const shapeAlignBtn = (
    align: 'left' | 'center' | 'right' | 'justify',
    title: string,
    icon: ReactNode,
  ) => (
    <button
      className={`rb-icon ${shapeTextActive.align === align ? 'active' : ''}`}
      disabled={!canEdit}
      data-tip={title}
      aria-label={title}
      onClick={() => shapeTextCommand.setAlign(align)}
    >
      {icon}
    </button>
  )

  return (
    <div
      className={`ribbon ${collapse.rootClass}`}
      ref={collapse.rootRef}
      onMouseDown={keepDocumentFocus}
    >
      <div
        className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}`}
        onDoubleClick={collapse.onTabsDoubleClick}
      >
        {!IS_MAC && (
          <div className="file-tab-wrap">
            <button
              className={`ribbon-tab ribbon-tab-file ${dropdown === 'file' ? 'open' : ''}`}
              onClick={() => setDropdown((v) => (v === 'file' ? null : 'file'))}
            >
              {t('ribbonTabFile')}
            </button>
            {dropdown === 'file' && (
              <div data-rb-panel="" className="file-menu">
                <button
                  onClick={() => {
                    setDropdown(null)
                    onOpen()
                  }}
                >
                  {t('ribbonOpen')} <span className="file-menu-key">Ctrl+O</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setDropdown(null)
                    onSave()
                  }}
                >
                  {t('ribbonSave')} <span className="file-menu-key">Ctrl+S</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setDropdown(null)
                    onSaveAs()
                  }}
                >
                  {t('ribbonSaveAs')} <span className="file-menu-key">Ctrl+Shift+S</span>
                </button>
              </div>
            )}
          </div>
        )}
        {quickActions}
        {TABS.filter((tabName) => tabName !== 'file').map((tabName) => (
          <button
            key={tabName}
            className={`ribbon-tab ${collapse.tabClass(tab === tabName)}`}
            data-tip={collapse.tabTip(tab === tabName)}
            onClick={() => {
              collapse.onTabPress(tab === tabName)
              lastRegularTab.current = tabName
              setTab(tabName)
              setDropdown(null)
            }}
          >
            {t(TAB_LABEL_KEYS[tabName])}
          </button>
        ))}
        {/* contextual tabs render as plain tabs appended to the row, like current Word */}
        {inTable &&
          TABLE_TABS.map((tableTab) => (
            <button
              key={tableTab}
              className={`ribbon-tab ${collapse.tabClass(tab === tableTab)}`}
              data-tip={collapse.tabTip(tab === tableTab)}
              onClick={() => {
                collapse.onTabPress(tab === tableTab)
                setTab(tableTab)
                setDropdown(null)
              }}
            >
              {t(TAB_LABEL_KEYS[tableTab])}
            </button>
          ))}
        {inImage &&
          IMAGE_TABS.map((imageTab) => (
            <button
              key={imageTab}
              className={`ribbon-tab ${collapse.tabClass(tab === imageTab)}`}
              data-tip={collapse.tabTip(tab === imageTab)}
              onClick={() => {
                collapse.onTabPress(tab === imageTab)
                setTab(imageTab)
                setDropdown(null)
              }}
            >
              {t(TAB_LABEL_KEYS[imageTab])}
            </button>
          ))}
        {inShape &&
          SHAPE_TABS.map((shapeTab) => (
            <button
              key={shapeTab}
              className={`ribbon-tab ${collapse.tabClass(tab === shapeTab)}`}
              data-tip={collapse.tabTip(tab === shapeTab)}
              onClick={() => {
                collapse.onTabPress(tab === shapeTab)
                setTab(shapeTab)
                setDropdown(null)
              }}
            >
              {t(TAB_LABEL_KEYS[shapeTab])}
            </button>
          ))}
        {inHf &&
          HF_TABS.map((hfTab) => (
            <button
              key={hfTab}
              className={`ribbon-tab ${tab === hfTab ? 'active' : ''}`}
              onClick={() => {
                collapse.onTabPress(tab === hfTab)
                setTab(hfTab)
                setDropdown(null)
              }}
            >
              {t(TAB_LABEL_KEYS[hfTab])}
            </button>
          ))}
        <span className="ribbon-tabs-spacer" />
        {trailingActions}
      </div>

      <div className="ribbon-body" data-ribbon-body="">
        {tab === 'headerFooter' && hfEditing ? (
          <HeaderFooterTab
            info={hfEditing}
            dropdown={dropdown}
            setDropdown={setDropdown}
            onAction={onHfAction}
          />
        ) : tab === 'shapeFormat' && inShape ? (
          <div className="table-ribbon-body">
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                {!shapeIsLine && (
                  <div className="rb-split-wrap">
                    <button
                      className="rb-big"
                      disabled={!canEdit}
                      data-tip={t('ribbonShapeFillTip')}
                      onClick={() => setDropdown((v) => (v === 'shapeFill' ? null : 'shapeFill'))}
                    >
                      <span className="rb-big-icon">
                        <IconShading />
                        <span
                          className="rb-color-bar"
                          style={{ background: fs.shapeFill ? `#${fs.shapeFill}` : 'transparent' }}
                        />
                      </span>
                      <span>{t('ribbonShapeFill')}</span>
                    </button>
                    {dropdown === 'shapeFill' && (
                      <RibbonColorPalette
                        current={fs.shapeFill}
                        noneLabel={t('ribbonNoFill')}
                        onPick={(hex) => {
                          setShapeStyle({ fill: hex })
                          setDropdown(null)
                        }}
                      />
                    )}
                  </div>
                )}
                <div className="rb-split-wrap">
                  <button
                    className="rb-big"
                    disabled={!canEdit}
                    data-tip={t('ribbonShapeOutlineTip')}
                    onClick={() =>
                      setDropdown((v) => (v === 'shapeOutline' ? null : 'shapeOutline'))
                    }
                  >
                    <span className="rb-big-icon">
                      <IconBorderAll />
                      <span
                        className="rb-color-bar"
                        style={{
                          background: fs.shapeBorderColor
                            ? `#${fs.shapeBorderColor}`
                            : 'transparent',
                        }}
                      />
                    </span>
                    <span>{t('ribbonShapeOutline')}</span>
                  </button>
                  {dropdown === 'shapeOutline' && (
                    <RibbonColorPalette
                      current={fs.shapeBorderColor}
                      noneLabel={t('ribbonNoOutline')}
                      onPick={(hex) => {
                        setShapeStyle({ borderColor: hex })
                        setDropdown(null)
                      }}
                    />
                  )}
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupShapeStyles')}</div>
            </div>
            {fs.shapeHasText && (
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <div className="rb-col">
                    <div className="rb-row">
                      {shapeMarkBtn('bold', shapeTextActive.bold, t('ribbonBoldTip'), <b>B</b>)}
                      {shapeMarkBtn(
                        'italic',
                        shapeTextActive.italic,
                        t('ribbonItalicTip'),
                        <i>I</i>,
                      )}
                      {shapeMarkBtn(
                        'underline',
                        shapeTextActive.underline,
                        t('ribbonUnderlineTip'),
                        <u>U</u>,
                      )}
                      <div className="rb-split-wrap">
                        <button
                          className="rb-icon rb-color-btn"
                          disabled={!canEdit}
                          data-tip={t('ribbonFontColor')}
                          aria-label={t('ribbonFontColor')}
                          onClick={() =>
                            setDropdown((v) => (v === 'shapeTextColor' ? null : 'shapeTextColor'))
                          }
                        >
                          <span className="rb-color-glyph rb-color-glyph-svg">
                            <IconFontColorA />
                            <span
                              className="rb-color-bar"
                              style={{
                                background: shapeTextActive.color
                                  ? `#${shapeTextActive.color}`
                                  : 'transparent',
                              }}
                            />
                          </span>
                        </button>
                        {dropdown === 'shapeTextColor' && (
                          <RibbonColorPalette
                            current={shapeTextActive.color}
                            noneLabel={t('ribbonAutomatic')}
                            onPick={(hex) => {
                              shapeTextCommand.setColor(hex)
                              setDropdown(null)
                            }}
                          />
                        )}
                      </div>
                    </div>
                    <div className="rb-row">
                      {shapeAlignBtn('left', t('ribbonAlignLeftTip'), <IconAlignLeft />)}
                      {shapeAlignBtn('center', t('ribbonAlignCenterTip'), <IconAlignCenter />)}
                      {shapeAlignBtn('right', t('ribbonAlignRightTip'), <IconAlignRight />)}
                      {shapeAlignBtn('justify', t('ribbonJustifyTip'), <IconAlignJustify />)}
                    </div>
                  </div>
                </div>
                <div className="ribbon-group-label">{t('ribbonGroupText')}</div>
              </div>
            )}
          </div>
        ) : tab === 'pictureFormat' && inImage ? (
          <div className="table-ribbon-body">
            {/* ---- Adjust: remove background / crop / replace picture ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <button
                  className="rb-big"
                  disabled={!canEdit}
                  data-tip={t('ribbonRemoveBgTip')}
                  onClick={() => setPictureDialog('cutout')}
                >
                  <span className="rb-big-icon">
                    <IconRemoveBg size={28} />
                  </span>
                  <span>{t('ribbonRemoveBg')}</span>
                </button>
                <button
                  className="rb-big"
                  disabled={!canEdit}
                  data-tip={t('ribbonCropTip')}
                  onClick={() => setPictureDialog('crop')}
                >
                  <span className="rb-big-icon">
                    <IconCrop size={28} />
                  </span>
                  <span>{t('ribbonCrop')}</span>
                </button>
                <button
                  className="rb-big"
                  disabled={!canEdit}
                  data-tip={t('ribbonReplacePictureTip')}
                  onClick={() => void replacePicture()}
                >
                  <span className="rb-big-icon">
                    <IconReplacePicture size={28} />
                  </span>
                  <span>{t('ribbonReplacePicture')}</span>
                </button>
                <div className="rb-col">
                  <button
                    className="rb-small"
                    disabled={!canEdit}
                    onClick={() => rotatePicture(90)}
                  >
                    <IconRotateRight size={18} />
                    <span>{t('ribbonRotateRight')}</span>
                  </button>
                  <button
                    className="rb-small"
                    disabled={!canEdit}
                    onClick={() => rotatePicture(-90)}
                  >
                    <IconRotateLeft size={18} />
                    <span>{t('ribbonRotateLeft')}</span>
                  </button>
                </div>
                <div className="rb-col">
                  <button
                    className={fs.imageFlipH ? 'rb-small active' : 'rb-small'}
                    disabled={!canEdit}
                    onClick={() => flipPicture('h')}
                  >
                    <IconFlipH size={18} />
                    <span>{t('ribbonFlipH')}</span>
                  </button>
                  <button
                    className={fs.imageFlipV ? 'rb-small active' : 'rb-small'}
                    disabled={!canEdit}
                    onClick={() => flipPicture('v')}
                  >
                    <IconFlipV size={18} />
                    <span>{t('ribbonFlipV')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupAdjust')}</div>
            </div>
            <div className="ribbon-sep" />
            {/* ---- Arrange: wrap text / align ---- */}
            <div className="table-tool-group">
              <div className="table-tool-row">
                <Dropdown
                  className="rb-wrap-dd"
                  disabled={!canEdit}
                  tip={t('ribbonWrapText')}
                  value={fs.imageWrap ?? ''}
                  options={WRAP_OPTIONS.map((opt) => ({
                    value: opt.value ?? '',
                    label: t(opt.labelKey),
                  }))}
                  onPick={(v) => {
                    if (!canEdit) return
                    editor
                      .chain()
                      .focus()
                      .updateAttributes('docProtected', { imageWrap: v || null })
                      .run()
                  }}
                />
              </div>
              <div className="table-tool-row">
                {(
                  [
                    ['left', <IconAlignLeft key="l" />, t('ribbonAlignLeftTip')],
                    ['center', <IconAlignCenter key="c" />, t('ribbonAlignCenterTip')],
                    ['right', <IconAlignRight key="r" />, t('ribbonAlignRightTip')],
                  ] as const
                ).map(([value, icon, label]) => (
                  <button
                    key={value}
                    className={
                      (fs.imageAlign ?? 'left') === value
                        ? 'table-tool-button active'
                        : 'table-tool-button'
                    }
                    disabled={!canEdit}
                    data-tip={label}
                    aria-label={label}
                    onClick={() => {
                      if (!canEdit) return
                      editor
                        .chain()
                        .focus()
                        .updateAttributes('docProtected', {
                          imageAlign: value === 'left' ? null : value,
                        })
                        .run()
                    }}
                  >
                    {icon}
                  </button>
                ))}
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupArrange')}</div>
            </div>
            <div className="ribbon-sep" />
            {/* ---- Size: height/width in the measurement unit, aspect lock, reset ---- */}
            <div className="table-tool-group">
              <div className="table-tool-row table-size-inputs" key={fs.imageKey ?? 'nosel'}>
                <label>
                  {t('ribbonPicHeight')}
                  <LengthInput
                    value={fs.imageHeightPx !== null ? pxToTwips(fs.imageHeightPx) : null}
                    min={PICTURE_TWIPS_MIN}
                    max={PICTURE_TWIPS_MAX}
                    ariaLabel={t('ribbonPicHeight')}
                    onCommit={(twips) => twips && setPictureSize('h', twips)}
                  />
                </label>
                <label>
                  {t('ribbonPicWidth')}
                  <LengthInput
                    value={fs.imageWidthPx !== null ? pxToTwips(fs.imageWidthPx) : null}
                    min={PICTURE_TWIPS_MIN}
                    max={PICTURE_TWIPS_MAX}
                    ariaLabel={t('ribbonPicWidth')}
                    onCommit={(twips) => twips && setPictureSize('w', twips)}
                  />
                </label>
              </div>
              <div className="table-tool-row">
                <label className="table-tool-check">
                  <input
                    type="checkbox"
                    checked={lockAspect}
                    onChange={(e) => setLockAspect(e.target.checked)}
                  />
                  {t('ribbonLockAspectRatio')}
                </label>
                <button data-tip={t('ribbonResetSizeTip')} onClick={() => void resetPictureSize()}>
                  {t('ribbonResetSize')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupSize')}</div>
            </div>
          </div>
        ) : tab === 'tableDesign' ? (
          <div className="table-ribbon-body">
            <div className="table-tool-group">
              <div className="table-style-gallery">
                <button
                  className="table-style-card"
                  data-tip={t('ribbonRemoveTableStyleTip')}
                  onClick={() => chain().updateAttributes('docTable', { tblStyleId: null }).run()}
                >
                  <span className="table-style-card-grid plain" />
                  <span>{t('ribbonNoStyle')}</span>
                </button>
                {TABLE_PRESETS.map((preset) => (
                  <button
                    key={preset.label}
                    className="table-style-card"
                    data-tip={t(preset.label)}
                    onClick={() => runTableCommand(applyTablePreset(preset))}
                  >
                    <span
                      className="table-style-card-grid"
                      style={{
                        borderColor: `#${preset.borderColor}`,
                        borderTopColor: `#${preset.headerFill ?? 'FFFFFF'}`,
                        background: `repeating-linear-gradient(to bottom,#${preset.band1Fill ?? 'FFFFFF'} 0 50%,#${preset.band2Fill ?? preset.band1Fill ?? 'FFFFFF'} 50% 100%)`,
                      }}
                    />
                    <span>{t(preset.label)}</span>
                  </button>
                ))}
                {[...(styles?.values() ?? [])]
                  .filter((info) => info.type === 'table' && info.styleId !== 'TableNormal')
                  .map((info) => (
                    <button
                      key={info.styleId}
                      className={
                        tableAttrs.tblStyleId === info.styleId
                          ? 'table-style-card active'
                          : 'table-style-card'
                      }
                      data-tip={t('ribbonApplyTableStyleTip', { name: info.name })}
                      onClick={() =>
                        chain().updateAttributes('docTable', { tblStyleId: info.styleId }).run()
                      }
                    >
                      <span
                        className="table-style-card-grid"
                        style={{
                          background: info.tableDisplay?.fill
                            ? `#${info.tableDisplay.fill}`
                            : undefined,
                          borderTopColor: info.tableDisplay?.firstRow?.fill
                            ? `#${info.tableDisplay.firstRow.fill}`
                            : undefined,
                        }}
                      />
                      <span>{info.name}</span>
                    </button>
                  ))}
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupTableStyles')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-style-options">
                {(
                  [
                    ['firstRow', 'ribbonTableFirstRow'],
                    ['lastRow', 'ribbonTableLastRow'],
                    ['bandedRows', 'ribbonTableBandedRows'],
                    ['firstColumn', 'ribbonTableFirstColumn'],
                    ['lastColumn', 'ribbonTableLastColumn'],
                    ['bandedColumns', 'ribbonTableBandedColumns'],
                  ] as Array<[keyof TableLook, StringKey]>
                ).map(([key, label]) => (
                  <button
                    key={key}
                    className={
                      tableLook[key]
                        ? 'table-option-toggle table-tool-button active'
                        : 'table-option-toggle table-tool-button'
                    }
                    aria-pressed={tableLook[key]}
                    onClick={() => runTableCommand(setTableLookOption(key, !tableLook[key]))}
                  >
                    <span className="table-option-check" aria-hidden="true" />
                    <span>{t(label)}</span>
                  </button>
                ))}
              </div>
              <div className="ribbon-group-label">{t('ribbonTableStyleOptions')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <div className="rb-split-wrap">
                  <button
                    className={`table-command-button${dropdown === 'tableShading' ? ' active' : ''}`}
                    data-tip={t('ribbonGroupShading')}
                    aria-expanded={dropdown === 'tableShading'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableShading' ? null : 'tableShading'))
                    }
                  >
                    <IconShading />
                    <span className="table-command-copy">
                      <span>{t('ribbonGroupShading')}</span>
                    </span>
                    <IconCaret size={12} />
                  </button>
                  {dropdown === 'tableShading' && (
                    <RibbonColorPalette
                      current={(editor.getAttributes('docTableCell').fill as string | null) ?? null}
                      noneLabel={t('ribbonNoColor')}
                      onPick={(hex) => {
                        runTableCommand(setCellAttr('fill', hex))
                        setDropdown(null)
                      }}
                    />
                  )}
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupShading')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group table-tool-borders">
              <div className="rb-split-wrap table-split">
                <button
                  className="table-split-main"
                  data-tip={t(TABLE_BORDER_ITEMS[lastBorderMode].label)}
                  onClick={() => applyCellBorders(lastBorderMode)}
                >
                  <LastBorderIcon />
                  <span>{t('ribbonGroupBorders')}</span>
                </button>
                <button
                  className={`table-split-caret${dropdown === 'tableBorders' ? ' active' : ''}`}
                  aria-label={t('ribbonGroupBorders')}
                  aria-expanded={dropdown === 'tableBorders'}
                  onClick={() =>
                    setDropdown((current) => (current === 'tableBorders' ? null : 'tableBorders'))
                  }
                >
                  <IconCaret size={12} />
                </button>
                {dropdown === 'tableBorders' && (
                  <div data-rb-panel="" className="layout-menu table-autofit-menu">
                    {TABLE_BORDER_MENU.map((mode) => {
                      const { label, Icon } = TABLE_BORDER_ITEMS[mode]
                      return (
                        <button
                          key={mode}
                          className={mode === lastBorderMode ? 'active' : ''}
                          onClick={() => {
                            applyCellBorders(mode)
                            setDropdown(null)
                          }}
                        >
                          <Icon size={17} />
                          <span>{t(label)}</span>
                        </button>
                      )
                    })}
                  </div>
                )}
              </div>
              <div className="table-border-opts">
                <Dropdown
                  tip={t('ribbonBorderStyle')}
                  ariaLabel={t('ribbonBorderStyle')}
                  className="table-border-style"
                  value={borderStyle}
                  options={BORDER_STYLE_OPTIONS}
                  onPick={setBorderStyle}
                />
                <Dropdown
                  tip={t('ribbonBorderWidth')}
                  ariaLabel={t('ribbonBorderWidth')}
                  value={String(borderSz)}
                  options={[4, 8, 12, 18, 24].map((sz) => ({
                    value: String(sz),
                    label: t('ribbonPtValue', { n: sz / 8 }),
                  }))}
                  onPick={(v) => setBorderSz(Number(v))}
                />
                <div className="rb-split-wrap">
                  <button
                    className={`table-pen-color${dropdown === 'tableBorderColor' ? ' active' : ''}`}
                    data-tip={t('ribbonBorderColor')}
                    aria-label={t('ribbonBorderColor')}
                    aria-expanded={dropdown === 'tableBorderColor'}
                    onClick={() =>
                      setDropdown((current) =>
                        current === 'tableBorderColor' ? null : 'tableBorderColor',
                      )
                    }
                  >
                    <span className="table-pen-swatch" style={{ background: `#${borderColor}` }} />
                    <IconCaret size={12} />
                  </button>
                  {dropdown === 'tableBorderColor' && (
                    <RibbonColorPalette
                      current={borderColor}
                      noneLabel={t('ribbonAutomatic')}
                      onPick={(hex) => {
                        setBorderColor(hex ?? '000000')
                        setDropdown(null)
                      }}
                    />
                  )}
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupBorders')}</div>
            </div>
          </div>
        ) : tab === 'tableLayout' ? (
          <div className="table-ribbon-body">
            <div className="table-tool-group table-tool-advanced">
              <div className="table-tool-stack table-tool-stack-three">
                <div className="rb-split-wrap">
                  <button
                    className={`table-command-row table-command-row-caret${dropdown === 'tableSelect' ? ' active' : ''}`}
                    aria-expanded={dropdown === 'tableSelect'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableSelect' ? null : 'tableSelect'))
                    }
                  >
                    <IconSelectCells size={17} />
                    <span>{t('ribbonSelect')}</span>
                    <IconCaret size={12} />
                  </button>
                  {dropdown === 'tableSelect' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      {TABLE_SELECT_ITEMS.map(([kind, label]) => (
                        <button
                          key={kind}
                          onClick={() => {
                            editor.view.focus()
                            enterSelectedTable(editor.state, editor.view.dispatch)
                            selectTablePart(kind)(editor.state, editor.view.dispatch)
                            setDropdown(null)
                          }}
                        >
                          <IconSelectCells size={17} />
                          <span>{t(label)}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <button
                  className={
                    tableGridlines
                      ? 'table-command-row table-tool-button active'
                      : 'table-command-row table-tool-button'
                  }
                  aria-pressed={!!tableGridlines}
                  data-tip={t('ribbonViewGridlinesTip')}
                  onClick={onToggleTableGridlines}
                >
                  <IconGridlines size={17} />
                  <span>{t('ribbonViewGridlines')}</span>
                </button>
                <button
                  className="table-command-row"
                  onClick={() => {
                    setDropdown(null)
                    onTableDialog?.('properties')
                  }}
                >
                  <IconTableProperties size={17} />
                  <span>{t('ribbonTableProperties')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonTableData')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <div className="rb-split-wrap">
                  <button
                    className={`table-command-button${dropdown === 'tableDelete' ? ' active' : ''}`}
                    aria-expanded={dropdown === 'tableDelete'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableDelete' ? null : 'tableDelete'))
                    }
                  >
                    <IconTableDelete />
                    <span className="table-command-copy">
                      <span>{t('ribbonTableDeleteMenu')}</span>
                    </span>
                    <IconCaret size={12} />
                  </button>
                  {dropdown === 'tableDelete' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      <button
                        onClick={() => {
                          setDropdown(null)
                          if (canEdit) onTableDialog?.('deleteCells')
                        }}
                      >
                        <IconDeleteCells size={17} />
                        <span>{t('ribbonDeleteCells')}</span>
                      </button>
                      {(
                        [
                          ['ribbonDeleteColumn', IconColDelete, deleteColumn],
                          ['ribbonDeleteRow', IconRowDelete, deleteRow],
                          ['ribbonDeleteTable', IconTableDelete, deleteTable],
                        ] as Array<[StringKey, (props: { size?: number }) => ReactNode, Command]>
                      ).map(([label, Icon, command]) => (
                        <button
                          key={label}
                          onClick={() => {
                            setDropdown(null)
                            runTableCommand(command)
                          }}
                        >
                          <Icon size={17} />
                          <span>{t(label)}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div className="table-tool-col">
                  <div className="table-tool-grid table-tool-grid-four">
                    <button onClick={() => runTableCommand(addRowBefore)}>
                      <IconRowInsertAbove />
                      {t('ribbonInsertAbove')}
                    </button>
                    <button onClick={() => runTableCommand(addRowAfter)}>
                      <IconRowInsertBelow />
                      {t('ribbonInsertBelow')}
                    </button>
                    <button onClick={() => runTableCommand(addColumnBefore)}>
                      <IconColInsertLeft />
                      {t('ribbonInsertLeft')}
                    </button>
                    <button onClick={() => runTableCommand(addColumnAfter)}>
                      <IconColInsertRight />
                      {t('ribbonInsertRight')}
                    </button>
                  </div>
                  <button
                    className="table-command-row"
                    onClick={() => canEdit && onTableDialog?.('insertCells')}
                  >
                    <IconInsertCells size={17} />
                    <span>{t('ribbonInsertCells')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupRowsCols')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <button disabled={!fs.canMergeCells} onClick={() => runTableCommand(mergeCells)}>
                  <IconMergeCells />
                  {t('ribbonMergeCells')}
                </button>
                <button onClick={() => canEdit && onTableDialog?.('splitCells')}>
                  <IconSplitCells />
                  {t('ribbonSplitCells')}
                </button>
                <button
                  data-tip={t('ribbonSplitTableTip')}
                  disabled={!fs.canSplitTable}
                  onClick={() => runTableCommand(splitTableAtSelection())}
                >
                  <IconSplitTable />
                  {t('ribbonSplitTable')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupMerge')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <div className="table-size-inputs" key={activeCellInfo?.key ?? 'nosel'}>
                  <label>
                    {t('ribbonRowHeight')}
                    <LengthInput
                      value={activeCellInfo?.heightTwips ?? null}
                      min={0}
                      max={maxRowHeightTwips}
                      allowEmpty
                      placeholder={t('ribbonAuto')}
                      ariaLabel={t('ribbonRowHeight')}
                      onCommit={(twips) => applyRowHeight(twips)}
                    />
                  </label>
                  <label>
                    {t('ribbonColumnWidth')}
                    <LengthInput
                      value={activeCellInfo?.widthTwips ?? null}
                      min={0}
                      max={pxToTwips(sectionContentWidthPx)}
                      placeholder={t('ribbonAuto')}
                      ariaLabel={t('ribbonColumnWidth')}
                      onCommit={(twips) => applyColumnWidth(twips)}
                    />
                  </label>
                </div>
                <div className="rb-split-wrap">
                  <button
                    className={`table-command-button table-autofit-button${dropdown === 'tableAutoFit' ? ' active' : ''}`}
                    data-tip={t('ribbonAutoFit')}
                    aria-expanded={dropdown === 'tableAutoFit'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableAutoFit' ? null : 'tableAutoFit'))
                    }
                  >
                    <IconAutoFit size={18} />
                    <span className="table-command-copy">
                      <span>{t('ribbonAutoFit')}</span>
                      <small>{t(tableAutoFitLabel)}</small>
                    </span>
                    <IconCaret size={12} />
                  </button>
                  {dropdown === 'tableAutoFit' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      {TABLE_AUTO_FIT_OPTIONS.map(([mode, label]) => (
                        <button
                          key={mode}
                          className={tableAutoFitMode === mode ? 'active' : ''}
                          onClick={() => {
                            runTableCommand(setTableAutoFit(mode, sectionContentWidthPx))
                            setDropdown(null)
                          }}
                        >
                          <IconAutoFit size={17} />
                          <span>{t(label)}</span>
                          <span className="table-menu-state" aria-hidden="true" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div className="table-tool-col">
                  <button
                    className="table-command-row"
                    onClick={() => {
                      if (!canEdit) return
                      enterSelectedTable(editor.state, editor.view.dispatch)
                      distributeRowsEvenly(editor.view)
                    }}
                  >
                    <IconDistributeRows size={17} />
                    <span>{t('ribbonDistributeRows')}</span>
                  </button>
                  <button
                    className="table-command-row"
                    onClick={() =>
                      runTableCommand(distributeSelectedColumns(sectionContentWidthPx))
                    }
                  >
                    <IconDistributeColumns size={17} />
                    <span>{t('ribbonDistributeColumns')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupCellSize')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <div className="cell-align-grid" role="group" aria-label={t('ribbonCellAlignment')}>
                  {CELL_ALIGN_GRID.map(([v, h, labelKey]) => (
                    <button
                      key={`${v}-${h}`}
                      className={
                        (activeCellInfo?.vAlign ?? 'top') === v && (fs.align ?? 'left') === h
                          ? 'active'
                          : ''
                      }
                      data-tip={t(labelKey)}
                      aria-label={t(labelKey)}
                      onClick={() => runTableCommand(setCellAlignment(v, h))}
                    >
                      <IconCellAlign v={v} h={h} size={16} />
                    </button>
                  ))}
                </div>
                <div className="table-tool-col">
                  <div className="rb-split-wrap">
                    <button
                      className={`table-command-row table-command-row-caret${dropdown === 'cellTextDirection' ? ' active' : ''}`}
                      data-tip={t('ribbonTextDirection')}
                      aria-expanded={dropdown === 'cellTextDirection'}
                      onClick={() =>
                        setDropdown((current) =>
                          current === 'cellTextDirection' ? null : 'cellTextDirection',
                        )
                      }
                    >
                      <IconTextDirection size={17} />
                      <span>{t('ribbonTextDirection')}</span>
                      <IconCaret size={12} />
                    </button>
                    {dropdown === 'cellTextDirection' && (
                      <div data-rb-panel="" className="layout-menu table-autofit-menu">
                        {CELL_TEXT_DIRECTIONS.map(([dir, labelKey]) => (
                          <button
                            key={dir}
                            className={(fs.cellTextDirection ?? 'lrTb') === dir ? 'active' : ''}
                            onClick={() => {
                              runTableCommand(setCellTextDirection(dir))
                              setDropdown(null)
                            }}
                          >
                            <IconTextDirection size={17} />
                            <span>{t(labelKey)}</span>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  <button
                    className="table-command-row"
                    onClick={() => {
                      setDropdown(null)
                      if (canEdit) onTableDialog?.('cellMargins')
                    }}
                  >
                    <IconCellMargins size={17} />
                    <span>{t('ribbonCellMarginsCmd')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupAlignment')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <button
                  className={tableHeader.active ? 'table-tool-button active' : 'table-tool-button'}
                  disabled={!tableHeader.enabled}
                  aria-pressed={tableHeader.active}
                  onClick={() => runTableCommand(toggleRepeatHeaderRows())}
                >
                  <IconRepeatHeader size={17} />
                  {t('ribbonRepeatHeaderRows')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupData')}</div>
            </div>
          </div>
        ) : tab === 'home' ? (
          <>
            {/* ---- Genspark AI (first slot: entry + one-click AI actions) ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <button
                  className={`rb-big ai-entry ${showAi ? 'active' : ''}`}
                  data-tip={t('aiOpenAssistant')}
                  onClick={onToggleAi}
                >
                  <span className="rb-big-icon">
                    <GensparkMark size={26} />
                  </span>
                  <span>Genspark AI</span>
                </button>
                <button
                  className="rb-big ai-entry"
                  disabled={docEmpty}
                  data-tip={t('aiSummarizeBtn')}
                  onClick={() => onAiPreset(t('aiSummarizePrompt'))}
                >
                  <span className="rb-big-icon">
                    <span className="ai-feature-icon" aria-hidden="true">
                      <svg
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path
                          d="M13.875 21H12H6.5C5.39543 21 4.5 20.1046 4.5 19V5C4.5 3.89543 5.39543 3 6.5 3H17.5C18.6046 3 19.5 3.89543 19.5 5V9V12V13"
                          strokeLinecap="round"
                        />
                        <path d="M8.00001 7H16" strokeLinecap="round" />
                        <path d="M8.00007 10.2032H14.0001" strokeLinecap="round" />
                        <path d="M8.00007 13.4062H12.0001" strokeLinecap="round" />
                        <path
                          d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </span>
                  </span>
                  <span>{t('aiSummarizeBtn')}</span>
                </button>
                <button
                  className="rb-big ai-entry"
                  disabled={docEmpty}
                  data-tip={t('aiPolishBtn')}
                  onClick={() =>
                    onAiPreset(
                      t(
                        editor.state.selection.empty ? 'aiPolishPrompt' : 'aiPolishSelectionPrompt',
                      ),
                    )
                  }
                >
                  <span className="rb-big-icon">
                    <span className="ai-feature-icon" aria-hidden="true">
                      <svg
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path
                          d="M5.00012 20.7481L8.80319 20.7482L21.7482 7.80317L17.945 4L5 16.945L5.00012 20.7481Z"
                          strokeLinejoin="round"
                        />
                        <path d="M15.1406 6.80469L18.9438 10.6079" />
                        <path
                          d="M8 3L8.22106 3.59745C8.51094 4.38087 8.65589 4.77259 8.94166 5.05833C9.22743 5.34409 9.61914 5.48903 10.4026 5.77893L11 6L10.4026 6.22107C9.61914 6.51097 9.22743 6.65592 8.94166 6.94167C8.65589 7.22741 8.51094 7.61913 8.22106 8.40255L8 9L7.77894 8.40255C7.48906 7.61913 7.34411 7.22741 7.05834 6.94167C6.77257 6.65592 6.38086 6.51097 5.59743 6.22107L5 6L5.59743 5.77893C6.38086 5.48903 6.77257 5.34409 7.05834 5.05833C7.34411 4.77259 7.48906 4.38087 7.77894 3.59745L8 3Z"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </span>
                  </span>
                  <span>{t('aiPolishBtn')}</span>
                </button>
                <button
                  className="rb-big ai-entry"
                  disabled={docEmpty}
                  data-tip={t('aiTidyBtn')}
                  onClick={() => onAiPreset(t('aiTidyPrompt'))}
                >
                  <span className="rb-big-icon">
                    <span className="ai-feature-icon" aria-hidden="true">
                      <svg
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="M4 5H20" strokeLinecap="round" />
                        <path d="M4 9H16" strokeLinecap="round" />
                        <path d="M4 13H11" strokeLinecap="round" />
                        <path d="M4 17H10" strokeLinecap="round" />
                        <path
                          d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </span>
                  </span>
                  <span>{t('aiTidyBtn')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">Genspark AI</div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Clipboard ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <div className="rb-split-wrap rb-big-split">
                  <button
                    className="rb-big"
                    disabled={!canEdit}
                    aria-label={t('ribbonPaste')}
                    onClick={() => pasteAs()}
                  >
                    <span className="rb-big-icon">
                      <IconPaste size={28} />
                    </span>
                  </button>
                  <button
                    className={`rb-big-split-caret ${dropdown === 'paste' ? 'active' : ''}`}
                    disabled={!canEdit}
                    aria-haspopup="menu"
                    aria-expanded={dropdown === 'paste'}
                    data-testid="ribbon-paste-menu"
                    onClick={() => setDropdown((v) => (v === 'paste' ? null : 'paste'))}
                  >
                    <span>{t('ribbonPaste')}</span>
                    <span className="rb-caret-inline">
                      <IconCaret />
                    </span>
                  </button>
                  {dropdown === 'paste' && (
                    <div data-rb-panel="" className="spacing-menu paste-menu" role="menu">
                      <button role="menuitem" onClick={() => pasteAs('source')}>
                        <IconPasteSource size={16} />
                        {t('appPasteKeepSource')}
                      </button>
                      <button role="menuitem" onClick={() => pasteAs('merge')}>
                        <IconPasteMerge size={16} />
                        {t('appPasteMergeFormat')}
                      </button>
                      <button role="menuitem" onClick={() => pasteAs('text')}>
                        <IconPasteText size={16} />
                        {t('appPasteTextOnly')}
                      </button>
                      <div className="spacing-menu-sep" />
                      <button
                        role="menuitem"
                        onClick={() => {
                          setDropdown(null)
                          setPasteSpecial(true)
                        }}
                      >
                        {t('ribbonPasteSpecial')}
                      </button>
                      {onPasteDefaults && (
                        <button
                          role="menuitem"
                          onClick={() => {
                            setDropdown(null)
                            onPasteDefaults()
                          }}
                        >
                          {t('ribbonSetDefaultPaste')}
                        </button>
                      )}
                    </div>
                  )}
                  {pasteSpecial &&
                    createPortal(
                      <PasteSpecialDialog editor={ed} onClose={() => setPasteSpecial(false)} />,
                      document.body,
                    )}
                </div>
                <div className="rb-col">
                  <button
                    className="rb-small"
                    disabled={!canEdit}
                    data-tip={t('ribbonCutTip')}
                    aria-label={t('ribbonCutTip')}
                    onClick={() => clipboard('cut')}
                  >
                    <IconCut />
                  </button>
                  <button
                    className="rb-small"
                    disabled={!hasDoc}
                    data-tip={t('ribbonCopyTip')}
                    aria-label={t('ribbonCopyTip')}
                    onClick={() => clipboard('copy')}
                  >
                    <IconCopy />
                  </button>
                  <button
                    className={`rb-small ${painter ? 'active' : ''}`}
                    disabled={!canEdit || !!sub}
                    aria-pressed={!!painter}
                    data-tip={painter ? t('ribbonPainterActiveTip') : t('ribbonPainterTip')}
                    aria-label={painter ? t('ribbonPainterActiveTip') : t('ribbonPainterTip')}
                    onClick={onPainterClick}
                    onDoubleClick={onPainterDoubleClick}
                  >
                    <IconFormatPainter />
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupClipboard')}</div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Font ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items rb-font-group">
                <div className="rb-row">
                  {/* Editable combobox (free-typed input + full preset dropdown): real
                      documents use fonts and sizes outside any fixed list (GB/T 9704
                      fonts, half sizes like 13.5pt) */}
                  <div className="rb-split-wrap" onPointerDown={markComboPress}>
                    <input
                      ref={fontInputRef}
                      className="rb-select rb-font-family"
                      disabled={!canEdit}
                      value={fontDraft ?? currentFont}
                      aria-label={t('ribbonFontFamilyTip')}
                      data-tip={t('ribbonFontFamilyTip')}
                      aria-autocomplete="both"
                      onChange={onFontTyped}
                      onKeyDown={onFontKeyDown}
                      // focusing the input relocates the DOM selection into it,
                      // hiding the document highlight — the decoration keeps the
                      // target text visibly selected, like Word (r119)
                      onFocus={(e) => {
                        e.currentTarget.select()
                        setInactiveSelectionShown(ed, true)
                      }}
                      onBlur={() => onComboBlur(resetFontBox, 'fontFamily')}
                    />
                    <button
                      className="rb-caret rb-combo-caret"
                      disabled={!canEdit}
                      data-tip={t('ribbonFontFamilyTip')}
                      aria-label={t('ribbonFontFamilyTip')}
                      onClick={() => {
                        if (dropdown !== 'fontFamily') loadSystemFonts()
                        setDropdown((v) => (v === 'fontFamily' ? null : 'fontFamily'))
                      }}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'fontFamily' && (
                      <div data-rb-panel="" className="spacing-menu rb-font-family-menu">
                        {(['b', 's'] as const).map((section) => {
                          const items = fontMenuItems.filter((it) => it.section === section)
                          if (items.length === 0) return null
                          return (
                            <div key={section}>
                              <div className="rb-menu-group-label">
                                {t(section === 'b' ? 'ribbonFontsCommon' : 'ribbonFontsSystem')}
                              </div>
                              {items.map((it) => (
                                <button
                                  key={it.key}
                                  data-font-key={it.key}
                                  className={`${it.name === currentFont ? 'active' : ''}${
                                    it.key === fontHighlight ? ' kbd-focus' : ''
                                  }`}
                                  // symbol fonts would render their own name as pictographs
                                  style={{
                                    fontFamily: isSymbolFontFamily(it.name)
                                      ? undefined
                                      : cssFontFamily(it.name),
                                  }}
                                  onClick={() => setFont(it.name)}
                                >
                                  {it.name}
                                </button>
                              ))}
                            </div>
                          )
                        })}
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap" onPointerDown={markComboPress}>
                    <input
                      className="rb-select rb-font-size"
                      type="text"
                      inputMode="decimal"
                      disabled={!canEdit}
                      value={
                        sizeDraft ?? (fs.fontSizeMixed ? '' : fontSizeLabel(currentSize, lang))
                      }
                      aria-label={t('ribbonFontSizeTip')}
                      data-tip={t('ribbonFontSizeTip')}
                      onChange={(e) => setSizeDraft(e.currentTarget.value)}
                      onKeyDown={(e) => {
                        if (e.nativeEvent.isComposing) return
                        if (e.key === 'Enter') {
                          e.preventDefault()
                          commitSize(e.currentTarget.value)
                        } else if (e.key === 'Escape') {
                          e.preventDefault()
                          e.stopPropagation()
                          setSizeDraft(null)
                          setDropdown(null)
                          chain().run()
                        }
                      }}
                      onFocus={(e) => {
                        e.currentTarget.select()
                        setInactiveSelectionShown(ed, true)
                      }}
                      onBlur={() => onComboBlur(() => setSizeDraft(null), 'fontSize')}
                    />
                    <button
                      className="rb-caret rb-combo-caret"
                      disabled={!canEdit}
                      data-tip={t('ribbonFontSizeTip')}
                      aria-label={t('ribbonFontSizeTip')}
                      onClick={() => setDropdown((v) => (v === 'fontSize' ? null : 'fontSize'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'fontSize' && (
                      <div data-rb-panel="" className="spacing-menu rb-font-size-menu">
                        {fontSizeOptions(lang).map((o) => (
                          <button
                            key={o.name}
                            className={
                              !fs.fontSizeMixed &&
                              Math.round(o.pt * 2) === Math.round(currentSize * 2)
                                ? 'active'
                                : ''
                            }
                            onClick={() => setTextStyle({ sizeHalfPoints: Math.round(o.pt * 2) })}
                          >
                            {o.name}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    data-tip={t('ribbonGrowFont')}
                    aria-label={t('ribbonGrowFont')}
                    onClick={() => stepFontSize(1)}
                  >
                    <IconGrowFont />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    data-tip={t('ribbonShrinkFont')}
                    aria-label={t('ribbonShrinkFont')}
                    onClick={() => stepFontSize(-1)}
                  >
                    <IconShrinkFont />
                  </button>
                  <span className="rb-mini-sep" />
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon"
                      disabled={!canEdit}
                      data-tip={t('ribbonChangeCase')}
                      onClick={() => setDropdown((v) => (v === 'case' ? null : 'case'))}
                    >
                      <IconChangeCase />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'case' && (
                      <div data-rb-panel="" className="spacing-menu case-menu">
                        <button onClick={() => changeCase('sentence')}>
                          {t('ribbonCaseSentence')}
                        </button>
                        <button onClick={() => changeCase('lower')}>{t('ribbonCaseLower')}</button>
                        <button onClick={() => changeCase('upper')}>{t('ribbonCaseUpper')}</button>
                        <button onClick={() => changeCase('title')}>{t('ribbonCaseTitle')}</button>
                        <button onClick={() => changeCase('toggle')}>
                          {t('ribbonCaseToggle')}
                        </button>
                        {WIDTH_CASE_LANGS.has(lang) && (
                          <>
                            <button onClick={() => changeCase('halfWidth')}>
                              {t('ribbonCaseHalfWidth')}
                            </button>
                            <button onClick={() => changeCase('fullWidth')}>
                              {t('ribbonCaseFullWidth')}
                            </button>
                          </>
                        )}
                      </div>
                    )}
                  </div>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    data-tip={t('ribbonClearFormatting')}
                    aria-label={t('ribbonClearFormatting')}
                    onClick={() => chain().unsetAllMarks().run()}
                  >
                    <IconClearFormat />
                  </button>
                </div>
                <div className="rb-row">
                  {markBtn('bold', fs.bold, t('ribbonBoldTip'), <b>B</b>)}
                  {markBtn('italic', fs.italic, t('ribbonItalicTip'), <i>I</i>)}
                  {markBtn('underline', fs.underline, t('ribbonUnderlineTip'), <u>U</u>)}
                  {markBtn('strike', fs.strike, t('ribbonStrikethrough'), <s>ab</s>)}
                  <button
                    className={`rb-icon ${fs.vertAlign === 'subscript' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonSubscript')}
                    onClick={() => toggleVertAlign('subscript')}
                  >
                    <IconSubscript />
                  </button>
                  <button
                    className={`rb-icon ${fs.vertAlign === 'superscript' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonSuperscript')}
                    onClick={() => toggleVertAlign('superscript')}
                  >
                    <IconSuperscript />
                  </button>
                  <span className="rb-mini-sep" />
                  {/* highlight: main button applies pen color, caret opens palette */}
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon rb-color-btn ${fs.highlight ? 'active' : ''}`}
                      disabled={!canEdit}
                      data-tip={t('ribbonTextHighlightColor')}
                      aria-label={t('ribbonTextHighlightColor')}
                      onClick={() =>
                        setTextStyle({
                          highlight: fs.highlight === penHighlight ? null : penHighlight,
                        })
                      }
                    >
                      <span className="rb-color-glyph rb-color-glyph-svg">
                        <IconHighlight />
                        <span
                          className="rb-color-bar"
                          style={{ background: HIGHLIGHT_CSS[penHighlight] }}
                        />
                      </span>
                    </button>
                    <button
                      className={`rb-caret rb-color-caret${dropdown === 'highlight' ? ' active' : ''}`}
                      disabled={!canEdit}
                      onClick={() => setDropdown((v) => (v === 'highlight' ? null : 'highlight'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'highlight' && (
                      <div
                        data-rb-panel=""
                        className="color-palette color-palette-highlight color-palette-highlight-word"
                      >
                        <div className="color-section-title color-highlight-title">
                          {t('ribbonHighlightColors')}
                        </div>
                        <div className="color-highlight-grid">
                          {HIGHLIGHTS.map((h) => (
                            <button
                              key={h}
                              className={`color-swatch color-highlight-swatch ${fs.highlight === h ? 'selected' : ''}`}
                              data-tip={h}
                              aria-label={h}
                              style={{ background: HIGHLIGHT_CSS[h] }}
                              onClick={() => {
                                setPenHighlight(h)
                                setTextStyle({ highlight: h })
                              }}
                            />
                          ))}
                        </div>
                        <button
                          className={`color-none color-highlight-none ${!fs.highlight ? 'selected' : ''}`}
                          onClick={() => setTextStyle({ highlight: null })}
                        >
                          {t('ribbonNoColor')}
                        </button>
                      </div>
                    )}
                  </div>
                  {/* font color: main button applies pen color, caret opens palette */}
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon rb-color-btn"
                      disabled={!canEdit}
                      data-tip={t('ribbonFontColor')}
                      onClick={() => setTextStyle({ color: penColor })}
                    >
                      <span className="rb-color-glyph rb-color-glyph-svg">
                        <IconFontColorA />
                        <span
                          className="rb-color-bar"
                          style={{ background: `#${penColor ?? '000000'}` }}
                        />
                      </span>
                    </button>
                    <button
                      className={`rb-caret rb-color-caret${dropdown === 'color' ? ' active' : ''}`}
                      disabled={!canEdit}
                      onClick={() => setDropdown((v) => (v === 'color' ? null : 'color'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'color' && (
                      <RibbonColorPalette
                        current={fs.textColor}
                        noneLabel={t('ribbonAutomatic')}
                        onPick={(hex) => {
                          if (!hex) {
                            setPenColor(null)
                            setTextStyle({ color: null })
                          } else {
                            setPenColor(hex)
                            setTextStyle({ color: hex })
                          }
                        }}
                      />
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Paragraph ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items rb-font-group">
                <div className="rb-row">
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.listBullet ? 'active' : ''}`}
                      disabled={!canEdit || !!sub}
                      data-tip={t('ribbonBullets')}
                      aria-label={t('ribbonBullets')}
                      onClick={() => toggleList('bullet')}
                    >
                      <IconBullets />
                    </button>
                    <button
                      className={`rb-caret${dropdown === 'bulletLib' ? ' active' : ''}`}
                      disabled={!canEdit || !!sub}
                      data-tip={t('ribbonBullets')}
                      aria-label={t('ribbonBullets')}
                      onClick={() => setDropdown((v) => (v === 'bulletLib' ? null : 'bulletLib'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'bulletLib' && (
                      <div data-rb-panel="" className="layout-menu list-gallery list-gallery-word">
                        {presetSection(
                          'ribbonRecentBullets',
                          recentListPresets?.bullets,
                          'bullets',
                          bulletPreview,
                        )}
                        <div className="list-gallery-title">{t('ribbonBulletLibTitle')}</div>
                        <button
                          className={`list-gallery-card list-gallery-none${!fs.listBullet && !fs.listOrdered ? ' selected' : ''}`}
                          onClick={() => {
                            clearList()
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonListNone')}
                        </button>
                        {BULLET_LIBRARY.map((glyph) => (
                          <button
                            key={glyph}
                            className="list-gallery-card list-gallery-glyph"
                            onClick={() => {
                              applyListPreset(bulletPresetLevels(glyph), 'bullets')
                              setDropdown(null)
                            }}
                          >
                            {glyph}
                          </button>
                        ))}
                        {presetSection(
                          'ribbonDocumentBullets',
                          documentListPresets?.bullets,
                          'bullets',
                          bulletPreview,
                        )}
                        {listLevelRow}
                        <button
                          className="list-gallery-define"
                          onClick={() => {
                            onListDialog?.('bullet')
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonDefineNewBullet')}…
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.listOrdered ? 'active' : ''}`}
                      disabled={!canEdit || !!sub}
                      data-tip={t('ribbonNumbering')}
                      aria-label={t('ribbonNumbering')}
                      onClick={() => toggleList('ordered')}
                    >
                      <IconNumbered />
                    </button>
                    <button
                      className={`rb-caret${dropdown === 'numberLib' ? ' active' : ''}`}
                      disabled={!canEdit || !!sub}
                      data-tip={t('ribbonNumbering')}
                      aria-label={t('ribbonNumbering')}
                      onClick={() => setDropdown((v) => (v === 'numberLib' ? null : 'numberLib'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'numberLib' && (
                      <div data-rb-panel="" className="layout-menu list-gallery list-gallery-word">
                        {presetSection(
                          'ribbonRecentNumbers',
                          recentListPresets?.numbers,
                          'numbers',
                          numberPreview,
                        )}
                        <div className="list-gallery-title">{t('ribbonNumberLibTitle')}</div>
                        <button
                          className={`list-gallery-card list-gallery-none${!fs.listBullet && !fs.listOrdered ? ' selected' : ''}`}
                          onClick={() => {
                            clearList()
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonListNone')}
                        </button>
                        {NUMBER_LIBRARY.map((n, i) => {
                          const levels = numberPresetLevels(n.numFmt, n.pattern)
                          return (
                            <button
                              key={i}
                              className="list-gallery-card list-gallery-preview"
                              onClick={() => {
                                applyListPreset(levels, 'numbers')
                                setDropdown(null)
                              }}
                            >
                              {numberPreview(levels)}
                            </button>
                          )
                        })}
                        {presetSection(
                          'ribbonDocumentNumbers',
                          documentListPresets?.numbers,
                          'numbers',
                          numberPreview,
                        )}
                        {listLevelRow}
                        <button
                          className="list-gallery-define"
                          onClick={() => {
                            onListDialog?.('number')
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonDefineNewNumber')}…
                        </button>
                        <button
                          className="list-gallery-define"
                          disabled={!fs.listOrdered}
                          onClick={() => {
                            onListDialog?.('value')
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonSetNumberingValue')}…
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon"
                      disabled={!canEdit || !!sub}
                      data-tip={t('ribbonMultilevelTip')}
                      aria-label={t('ribbonMultilevelTip')}
                      onClick={() => setDropdown((v) => (v === 'multiLib' ? null : 'multiLib'))}
                    >
                      <IconMultilevel />
                    </button>
                    {dropdown === 'multiLib' && (
                      <div data-rb-panel="" className="layout-menu list-gallery list-gallery-multi">
                        {presetSection(
                          'ribbonRecentLists',
                          recentListPresets?.multi,
                          'multi',
                          multiPreview,
                        )}
                        <div className="list-gallery-title">{t('ribbonListLibraryTitle')}</div>
                        {MULTILEVEL_LIBRARY.map((levels, i) => (
                          <button
                            key={i}
                            className="list-gallery-card list-gallery-card-multi"
                            onClick={() => {
                              applyListPreset(levels, 'multi')
                              setDropdown(null)
                            }}
                          >
                            {multiPreview(levels)}
                          </button>
                        ))}
                        {presetSection(
                          'ribbonDocumentLists',
                          documentListPresets?.multi,
                          'multi',
                          multiPreview,
                        )}
                        {listLevelRow}
                        <button
                          className="list-gallery-define"
                          onClick={() => {
                            onListDialog?.('multi')
                            setDropdown(null)
                          }}
                        >
                          {t('ribbonDefineNewList')}…
                        </button>
                      </div>
                    )}
                  </div>
                  <span className="rb-mini-sep" />
                  <button
                    className="rb-icon"
                    disabled={!canEdit || !!sub}
                    data-tip={t('ribbonDecreaseIndent')}
                    aria-label={t('ribbonDecreaseIndent')}
                    onClick={() => changeIndent(-1)}
                  >
                    <IconIndentDec />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!canEdit || !!sub}
                    data-tip={t('ribbonIncreaseIndent')}
                    aria-label={t('ribbonIncreaseIndent')}
                    onClick={() => changeIndent(1)}
                  >
                    <IconIndentInc />
                  </button>
                  <span className="rb-mini-sep" />
                  <button
                    className={`rb-icon ${showMarks ? 'active' : ''}`}
                    disabled={!hasDoc}
                    data-tip={t('ribbonShowMarks')}
                    aria-label={t('ribbonShowMarks')}
                    onClick={() => onShowMarks(!showMarks)}
                  >
                    <IconPilcrow />
                  </button>
                </div>
                <div className="rb-row">
                  <button
                    className={`rb-icon ${activeAlign === 'left' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonAlignLeftTip')}
                    aria-label={t('ribbonAlignLeftTip')}
                    onClick={() => setSelectionAlign(ed, 'left')}
                  >
                    <IconAlignLeft />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'center' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonAlignCenterTip')}
                    aria-label={t('ribbonAlignCenterTip')}
                    onClick={() => setSelectionAlign(ed, 'center')}
                  >
                    <IconAlignCenter />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'right' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonAlignRightTip')}
                    aria-label={t('ribbonAlignRightTip')}
                    onClick={() => setSelectionAlign(ed, 'right')}
                  >
                    <IconAlignRight />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'justify' ? 'active' : ''}`}
                    disabled={!canEdit}
                    data-tip={t('ribbonJustifyTip')}
                    aria-label={t('ribbonJustifyTip')}
                    onClick={() => setSelectionAlign(ed, 'justify')}
                  >
                    <IconAlignJustify />
                  </button>
                  {showDirection && (
                    <>
                      <span className="rb-mini-sep" />
                      <button
                        className={`rb-icon ${!fs.bidi ? 'active' : ''}`}
                        disabled={!canEdit || !!sub}
                        data-tip={t('ribbonDirLtrTip')}
                        aria-label={t('ribbonDirLtrTip')}
                        onClick={() => setParagraphDirection(editor, 'ltr')}
                      >
                        <IconDirLtr />
                      </button>
                      <button
                        className={`rb-icon ${fs.bidi ? 'active' : ''}`}
                        disabled={!canEdit || !!sub}
                        data-tip={t('ribbonDirRtlTip')}
                        aria-label={t('ribbonDirRtlTip')}
                        onClick={() => setParagraphDirection(editor, 'rtl')}
                      >
                        <IconDirRtl />
                      </button>
                    </>
                  )}
                  <span className="rb-mini-sep" />
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${activeSpacing ? 'active' : ''}`}
                      disabled={!canEdit}
                      data-tip={t('ribbonLineSpacing')}
                      aria-label={t('ribbonLineSpacing')}
                      onClick={() => setDropdown((v) => (v === 'spacing' ? null : 'spacing'))}
                    >
                      <IconLineSpacing />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'spacing' && (
                      <div data-rb-panel="" className="spacing-menu">
                        {LINE_SPACINGS.map((s) => (
                          <button
                            key={s}
                            className={activeSpacing === s ? 'active' : ''}
                            // presets are multiples: clear any atLeast/exact rule so they take effect
                            onClick={() =>
                              setParaAttr({ lineSpacing: s, lineRule: null, lineRawTwips: null })
                            }
                          >
                            {s.toFixed(2).replace(/0+$/, '').replace(/\.$/, '.0')}
                          </button>
                        ))}
                        <div className="spacing-menu-sep" />
                        <button onClick={() => toggleParaSpace('spaceBefore')}>
                          {t(
                            fs.spaceBefore > 0 ? 'ribbonRemoveSpaceBefore' : 'ribbonAddSpaceBefore',
                          )}
                        </button>
                        <button onClick={() => toggleParaSpace('spaceAfter')}>
                          {t(fs.spaceAfter > 0 ? 'ribbonRemoveSpaceAfter' : 'ribbonAddSpaceAfter')}
                        </button>
                        {onParagraphDialog && (
                          <button
                            onClick={() => {
                              setDropdown(null)
                              onParagraphDialog()
                            }}
                          >
                            {t('ribbonLineSpacingOptions')}
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.shadingFill ? 'active' : ''}`}
                      disabled={!canEdit}
                      data-tip={t('ribbonParagraphShading')}
                      aria-label={t('ribbonParagraphShading')}
                      onClick={() => setDropdown((v) => (v === 'shading' ? null : 'shading'))}
                    >
                      <IconShading />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'shading' && (
                      <RibbonColorPalette
                        current={fs.shadingFill}
                        noneLabel={t('ribbonNoColor')}
                        onPick={(hex) => {
                          setParaAttr({ shadingFill: hex })
                          setDropdown(null)
                        }}
                      />
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.paraBorders ? 'active' : ''}`}
                      disabled={!canEdit}
                      data-tip={t('ribbonParagraphBorders')}
                      aria-label={t('ribbonParagraphBorders')}
                      onClick={() => setDropdown((v) => (v === 'borders' ? null : 'borders'))}
                    >
                      <IconBorderAll />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'borders' && (
                      <div data-rb-panel="" className="spacing-menu borders-menu">
                        <button onClick={() => setParaAttr({ borders: 'b' })}>
                          {t('ribbonBorderBottom')}
                        </button>
                        <button onClick={() => setParaAttr({ borders: 't' })}>
                          {t('ribbonBorderTop')}
                        </button>
                        <button onClick={() => setParaAttr({ borders: 'l' })}>
                          {t('ribbonBorderLeft')}
                        </button>
                        <button onClick={() => setParaAttr({ borders: 'r' })}>
                          {t('ribbonBorderRight')}
                        </button>
                        <button onClick={() => setParaAttr({ borders: 'tblr' })}>
                          {t('ribbonBorderBox')}
                        </button>
                        <button onClick={() => setParaAttr({ borders: null })}>
                          {t('ribbonNoBorders')}
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupParagraph')}</div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Styles ---- */}
            <div className="ribbon-group ribbon-group-styles">
              <div className="ribbon-group-items rb-split-wrap style-gallery-wrap">
                <div className="style-gallery" ref={styleGalleryRef}>
                  {renderStyleCards(false)}
                </div>
                {/* clipped cards stay reachable through the expander grid */}
                {styleGalleryOverflow && (
                  <button
                    className="style-gallery-more"
                    data-tip={t('ribbonMoreStyles')}
                    aria-label={t('ribbonMoreStyles')}
                    aria-expanded={dropdown === 'styleGallery'}
                    onClick={() =>
                      setDropdown((v) => (v === 'styleGallery' ? null : 'styleGallery'))
                    }
                  >
                    <IconCaret />
                  </button>
                )}
                {dropdown === 'styleGallery' && (
                  <div data-rb-panel="" className="style-gallery-menu">
                    {renderStyleCards(true)}
                  </div>
                )}
                <button
                  className={`rb-big rb-styles-pane${showStylesPane ? ' active' : ''}`}
                  disabled={!hasDoc}
                  aria-pressed={!!showStylesPane}
                  data-tip={t('ribbonStylesPaneTip')}
                  onClick={() => onShowStylesPane?.(!showStylesPane)}
                >
                  <span className="rb-big-icon">
                    <IconStylesPane />
                  </span>
                  <span>{t('ribbonStylesPane')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupStyles')}</div>
            </div>

            {/* Editing group is Word for Windows only; Word for Mac keeps Find in the search field */}
            {!IS_MAC && (
              <>
                <div className="ribbon-sep" />
                <div className="ribbon-group">
                  <div className="ribbon-group-items rb-col rb-editing">
                    <button
                      className="rb-small"
                      disabled={!hasDoc}
                      data-tip={t('ribbonFindTip')}
                      onClick={() => onFind?.()}
                    >
                      <IconSearch />
                      {t('ribbonFind')}
                    </button>
                    <button
                      className="rb-small"
                      disabled={!canEdit}
                      data-tip={t('ribbonReplaceTip')}
                      onClick={() => onReplace?.()}
                    >
                      <IconReplace />
                      {t('ribbonReplace')}
                    </button>
                    <button
                      className="rb-small"
                      disabled={!hasDoc}
                      data-tip={t('ribbonSelectAllTip')}
                      onClick={() => ed.chain().focus().selectAll().run()}
                    >
                      <IconSelectAll />
                      {t('ribbonSelectAll')}
                    </button>
                  </div>
                  <div className="ribbon-group-label">{t('ribbonGroupEditing')}</div>
                </div>
              </>
            )}
          </>
        ) : tab === 'draw' ? (
          <DrawTab
            hasDoc={hasDoc}
            tool={inkTool}
            onTool={onInkTool}
            pen={inkPen}
            onPen={onInkPen}
            highlighter={inkHighlighter}
            onHighlighter={onInkHighlighter}
            annotationCount={inkCount}
            onClearAll={onInkClearAll}
          />
        ) : tab === 'insert' ? (
          <InsertTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            header={header}
            onHeader={onHeader}
            onPageNumFormat={onPageNumFormat}
            onInsertField={onInsertField}
            footer={footer}
            onFooter={onFooter}
            onHfEdit={onHfEdit}
            canComment={canComment}
            onNewComment={onNewComment}
            isProtected={isProtected}
            commentsAllowed={commentsAllowed}
            onTableInserted={() => setTab('tableDesign')}
          />
        ) : tab === 'design' ? (
          <DesignTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            pageColor={pageColor}
            onPageColor={onPageColor}
            section={section}
            onSection={onSection}
            watermark={watermark}
            onWatermark={onWatermark}
            themeFonts={themeFonts}
            onThemeFonts={onThemeFonts}
            onThemeColors={onThemeColors}
          />
        ) : tab === 'layout' ? (
          <LayoutTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            section={section}
            onSection={onSection}
            activeSection={activeSection}
            onInsertSectionBreak={onInsertSectionBreak}
            onPaperSizeAll={onPaperSizeAll}
            mirrorMargins={mirrorMargins}
            onMirrorMargins={onMirrorMargins}
          />
        ) : tab === 'references' ? (
          <ReferencesTab
            editor={editor}
            hasDoc={canEdit}
            blocks={blocks}
            dropdown={dropdown}
            setDropdown={setDropdown}
            onInsertNote={onInsertNote}
            sources={sources}
            zoteroNoteFields={zoteroNoteFields}
            onAddSource={onAddSource}
            headingPages={headingPages}
          />
        ) : tab === 'review' ? (
          <ReviewTab
            editor={editor}
            hasDoc={hasDoc}
            dropdown={dropdown}
            setDropdown={setDropdown}
            onAiPreset={onAiPreset}
            commentCount={commentCount}
            openCommentCount={openCommentCount}
            resolvedCommentCount={resolvedCommentCount}
            onShowComments={onShowComments}
            canComment={canComment}
            onNewComment={onNewComment}
            commentAtCaret={commentAtCaret}
            onDeleteComment={onDeleteComment}
            onDeleteAllComments={onDeleteAllComments}
            onGotoComment={onGotoComment}
            trackChanges={trackChanges}
            onTrackChanges={onTrackChanges}
            spellcheck={spellcheck}
            onSpellcheck={onSpellcheck}
            revisionDisplay={revisionDisplay}
            onRevisionDisplay={onRevisionDisplay}
            revisionCount={revisionCount}
            onAcceptRevision={onAcceptRevision}
            onRejectRevision={onRejectRevision}
            onGotoRevision={onGotoRevision}
            isProtected={isProtected}
            commentsAllowed={commentsAllowed}
            trackChangesForced={trackChangesForced}
            protectActive={protectActive}
            onProtectDoc={onProtectDoc}
            onCompare={onCompare}
          />
        ) : (
          <ViewTab
            hasDoc={hasDoc}
            filePath={filePath}
            zoom={zoom}
            onZoom={onZoom}
            onZoomFit={onZoomFit}
            onZoomDialog={onZoomDialog}
            showAi={showAi}
            onToggleAi={onToggleAi}
            darkPage={darkPage}
            onDarkPage={onDarkPage}
            showRuler={showRuler}
            onShowRuler={onShowRuler}
            showNav={showNav}
            onShowNav={onShowNav}
            viewMode={viewMode}
            onViewMode={onViewMode}
            readMode={readMode}
            onReadMode={onReadMode}
            showGrid={showGrid}
            onShowGrid={onShowGrid}
            splitView={splitView}
            onSplitView={onSplitView}
            onPagePreview={onPagePreview}
          />
        )}
      </div>

      {pictureDialog === 'cutout' && imageDataUrl && (
        <CutoutDialog
          dataUrl={imageDataUrl}
          onApply={(png) => {
            setPictureDialog(null)
            void applyPictureBytes(png)
          }}
          onCancel={() => setPictureDialog(null)}
        />
      )}
      {pictureDialog === 'crop' && imageDataUrl && (
        <CropDialog
          dataUrl={imageDataUrl}
          onApply={(cropped) => {
            setPictureDialog(null)
            void applyPictureBytes(cropped)
          }}
          onCancel={() => setPictureDialog(null)}
        />
      )}
    </div>
  )
}

// memo + shallow-stable formatState/callback props: caret moves that change no
// displayed format skip re-rendering the whole ribbon
export const Ribbon = memo(RibbonInner)
