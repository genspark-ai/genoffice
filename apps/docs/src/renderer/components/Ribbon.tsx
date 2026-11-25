import { memo, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { wrapBigCaptions } from './ribbon-caption-wrap'
import { fitRibbon, styleCardRoom } from './ribbon-fit'
import { RbGroupMenu } from './ribbon-group-menu'
import type { StyleSetSource } from '../style-set-previews'
import type { SavedTheme } from '../saved-themes'
import { ArrangeGroup } from './ribbon-arrange-group'
import type { Lang } from '@genoffice/i18n'

import type {
  ChangeEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactNode,
} from 'react'
import type { ChainedCommands, Editor } from '@tiptap/core'
import { NodeSelection, type Command } from '@tiptap/pm/state'
import type { Mark, Node as PMNode } from '@tiptap/pm/model'
import { selectedRect, setCellAttr } from '@tiptap/pm/tables'
import type {
  CustomNumberingLevel,
  DocDefaults,
  HeaderFooter,
  HfImage,
  Run,
  SectionSettings,
  CitationStyleId,
  SourceInfo,
  StyleInfo,
  TableAutoFitMode,
  TableLook,
  TextboxDisplay,
  TextboxParaDisplay,
  ThemeColors,
  ThemeFonts,
  PictureWatermarkSpec,
} from '@genoffice/docx-engine'
import {
  PAINTER_BLOCK_EXTRA,
  PAINTER_MARK_TYPES,
  PAINTER_PARA_KEYS,
  applyPainterFormat,
  painterWordRangeAt,
  type PainterFormat,
  type PainterState,
} from './ribbon-painter'
import {
  Dropdown,
  isSymbolFontFamily,
  useDismissablePopover,
  useRibbonCollapse,
  useScrollWindow,
} from '@genoffice/ui'
import { HIGHLIGHT_CSS } from '../editor/extensions'
import type { CaseMode } from '../editor/case-transform'
import { isRtlUiLang, setSelectionAlign } from '../editor/direction'
import { selectedBoxIndex } from '../editor/selected-box'
import { HORIZONTAL, TEXTBOX_DIRECTIONS } from '../editor/textbox-direction'
import { setInactiveSelectionShown } from '../editor/inactive-selection'
import type { PasteMode } from '../editor/paste-options'
import { formatNumber } from '../editor/numbering'
import { getActiveSubEditor } from '../editor/active-editor'
import { originalPictureSizeOf } from '../editor/picture-bytes'
import { selectedGraphicFill } from '../editor/graphics-fill'
import { listLibrary, previewLevelText, type ListPresetKind } from '../list-presets'
import type { InkTool } from '../editor/ink'
import type { InkPreset } from '../editor/ink-presets'
import {
  RbCommand,
  RbMenuCommand,
  RibbonCommandContext,
  useRibbonCommand,
  useRunCommand,
  type RibbonCommandSource,
} from './ribbon-command'
import { RbTextEffects } from './ribbon-text-effects'
import type { RibbonFormatState } from './ribbon-format-state'
import { setSelectedColumnWidth } from '../editor/table-sizing'
import {
  enterSelectedTable,
  inTableOrSelected,
  tableCellsSelection,
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
import { repeatHeaderState } from '../editor/table-properties'
import { useI18n, type StringKey, type TFunc } from '../i18n/locale'
import type { ActionCommandId } from '../../shared/commands'
import type { EditTarget } from '../commands/edit-target'
import { NONE, type ListPresetPayload } from '../commands/menu-commands'
import { clipboardOff } from '../commands/menu-state'
import type { MergeDataSource } from '../../shared/ipc'
import type { MergeRange } from '../mail-merge/range'
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
import { HeaderFooterTab, type HfEditingInfo } from './ribbon-hf-tab'
import { fontFamiliesFor, systemFamiliesBesidesCandidates } from '../font-list'
import {
  fontSizeLabel,
  fontSizeOptions,
  parseFontSize,
  stepFontSize as stepSizeInList,
} from '../font-sizes'
import { useSystemFontFamilies } from '../system-fonts'
import { cssFontFamily } from '../text-layout/font-rules'
import type { RevisionFilter } from '../editor/revision-filters'
import {
  DesignTab,
  DrawTab,
  InsertTab,
  LayoutTab,
  MailingsTab,
  ReferencesTab,
  ReviewTab,
  ViewTab,
  type BalloonMode,
  type RevisionDisplayMode,
  type ViewMode,
  setParaAttrs,
  BIG,
} from './ribbon-tabs'
import { ZoteroTab } from './ribbon-zotero-tab'
import type { CommentsMode } from './comments-mode'
import {
  GensparkMark,
  IconAlignCenter,
  IconAlignJustify,
  IconAlignLeft,
  IconAlignRight,
  IconAutoFit,
  IconBorderAll,
  IconBorderBottom,
  IconBorderDiagonalDown,
  IconBorderDiagonalUp,
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
  IconConvert,
  IconColDelete,
  IconColInsertLeft,
  IconColInsertRight,
  IconMultilevel,
  IconClearFormat,
  IconPhoneticGuide,
  IconCharBorder,
  IconCharShading,
  IconEncloseChars,
  IconAsianLayout,
  IconSort,
  IconAlignDistributed,
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
  IconPicture,
  IconPilcrow,
  IconRemoveBg,
  IconReplacePicture,
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
  IconStyles,
  IconStylesPane,
  IconSelectAll,
  IconComments,
} from './icons'
interface RibbonProps {
  /** App keyboard shortcuts reuse ribbon closures through here (font-size stepping keeps its coalescing) */
  actionsRef?: React.MutableRefObject<{
    stepFontSize?: (dir: 1 | -1) => void
    nudgeFontSize?: (dir: 1 | -1) => void
    toggleRibbon?: () => void
  }>
  /** told whenever the command band collapses or comes back, for View > Ribbon's check mark */
  onRibbonCollapsed?: (collapsed: boolean) => void
  /** the document asks to be opened read-only (Review > Always Open Read-Only shows pressed) */
  readOnlyRecommended?: boolean
  /** Mailings > Preview Results is on (its button shows pressed) */
  mergePreview?: boolean
  /** a Docs window stands outside the tab strip: Merge All Windows has one to bring back */
  windowsDetached?: boolean
  /** Quick-access area on the tab row's left (save/undo-redo/autosave), matching the WPS/Office QAT */
  quickActions?: React.ReactNode
  /** Right side of the tab row (file name, etc.) */
  trailingActions?: React.ReactNode
  editor: Editor
  /** shallow-stable snapshot of every editor-state read shown in the ribbon (memo invalidation key) */
  formatState: RibbonFormatState
  hasDoc: boolean
  /** formats already used in the document, for the "Document …" gallery sections */
  documentListPresets?: Record<ListPresetKind, CustomNumberingLevel[][]>
  /** recently picked presets (localStorage), refreshed by the host after each pick */
  recentListPresets?: Record<ListPresetKind, CustomNumberingLevel[][]>
  /** document character styles, from ParsedDoc.styles (type === 'character') */
  styles?: Map<string, StyleInfo>
  /** document-wide text defaults from styles.xml */
  docDefaults?: DocDefaults
  /** Word's table dialogs live in App so the right-click menu and the Table menu share them */
  onTableDialog?: (kind: TableDialogKind, tab?: TablePropertiesTab) => void
  /** screen-only dashed gridlines on borderless tables (Table Design ▸ View Gridlines) */
  tableGridlines?: boolean
  /** what Cut and Copy would act on, as the menu bar sees it */
  editTarget: EditTarget
  /** runs a command of the shared command table, as the menu bar does */
  onCommand: (id: ActionCommandId, payload?: string) => void
  onOpen: () => void
  onSave: () => void
  onSaveAs: () => void
  /** File ▸ Name with AI…: ask the model for a name and offer it in the dialog */
  onNameWithAi: () => void
  showAi: boolean
  section: SectionSettings | null
  onSection: (next: SectionSettings) => void
  /** Multi-section documents: index of the cursor's section (0-based); null for single-section */
  activeSection: number | null
  /** page border for every section (Page Border dialog > Whole document); null removes it */
  onPageBorderAll: (props: SectionSettings['pageBorderProps'] | null) => void
  /** paper size for every section (More Paper Sizes… > Whole document); portrait dimensions */
  onPaperSizeAll: (
    portraitW: number,
    portraitH: number,
    orientation: SectionSettings['orientation'],
  ) => void
  mirrorMargins: boolean
  pageColor: string | null
  /** Design → Watermark / Themes */
  watermark: string | null
  watermarkPicture: PictureWatermarkSpec | null
  watermarkDocImage: HfImage | null
  themeFonts: ThemeFonts | null
  /** what Design's style set cards are resolved against */
  styleSetSource?: StyleSetSource | null
  /** the theme files of the themes folder: Design > Themes > Custom */
  savedThemes?: readonly SavedTheme[]
  themeColors: ThemeColors | null
  /** Draw → pen / highlighter / eraser */
  inkTool: InkTool
  onInkTool: (tool: InkTool) => void
  inkPresets: InkPreset[]
  inkActiveId: string
  onInkPresets: (presets: InkPreset[]) => void
  onInkActivate: (id: string) => void
  inkCount: number
  sources: SourceInfo[]
  /** the document's citation style; null when its sources name one the picker does not offer */
  citationStyle?: CitationStyleId | null
  /** footnotes/endnotes hold Zotero citation fields the bridge cannot see yet */
  zoteroNoteFields?: boolean
  showZotero?: boolean
  onAddSource: (source: SourceInfo) => void
  zoom: number
  darkPage: boolean
  /** external request (e.g. native menu Page Setup) to switch to a specific tab */
  tabRequest?: { tab: string; nonce: number; dropdown?: string } | null
  header: HeaderFooter | null
  footer: HeaderFooter | null
  /** a header/footer strip is being edited: the contextual Header & Footer tab shows */
  hfEditing: HfEditingInfo | null
  showMarks: boolean
  showRuler: boolean
  showNav: boolean
  /** Home > Styles > Styles Pane (Word for Mac) */
  showStylesPane?: boolean
  commentCount: number
  /** unresolved root comments (drives the AI resolve-comments action) */
  openCommentCount: number
  resolvedCommentCount: number
  /** Word-style top-right cluster: Show Comments toggle with its Contextual / List caret */
  showComments: boolean
  onToggleComments: () => void
  commentsMode: CommentsMode
  onCommentsMode: (mode: CommentsMode) => void
  /** Review → comments / revisions / compare / protection */
  commentAtCaret: boolean
  trackChanges: boolean
  onTrackChanges: (on: boolean) => void
  /** native check-as-you-type spellcheck (red squiggle) */
  spellcheck: boolean
  revisionDisplay: RevisionDisplayMode
  balloonMode: BalloonMode
  revisionCount: number
  revisionFilter: RevisionFilter
  revisionAuthors: string[]
  showReviewPane: boolean
  isProtected: boolean
  /** comments restriction: adding comments stays allowed although the body is read-only */
  commentsAllowed: boolean
  /** trackedChanges restriction: the recorder is forced on (toggle and accept/reject disabled) */
  trackChangesForced: boolean
  /** any protection is configured (highlights the Protect Document button) */
  protectActive: boolean
  /** Mailings: the data source the user picked, null when there is none yet */
  mergeSource: MergeDataSource | null
  /** Mailings: the record the preview shows */
  mergeRecordIndex: number
  onMergeRecordIndex: (index: number | ((index: number) => number)) => void
  mergeRange: MergeRange
  onMergeRange: (range: MergeRange) => void
  viewMode: ViewMode
  readMode: boolean
  onReadMode: (v: boolean) => void
  showGrid: boolean
  splitView: boolean
}

// Word for Mac has no File ribbon tab: file actions live in the native menu
// bar (which we provide). Windows Word does have one, so keep it there.
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
// Home's last rungs are Word for Mac's. Off macOS Home also carries the Editing group and
// already overflows the default window, where they would fold Paragraph away.
const HOME_FOLDS = IS_MAC
// Word offers Half-width / Full-width under Change Case only on CJK UIs
const WIDTH_CASE_LANGS = new Set<Lang>(['zh', 'zh-TW', 'ja'])
/** shell tab mode: the tab strip above owns traffic lights / caption buttons */
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'

const TABS = (
  IS_MAC
    ? ['home', 'insert', 'draw', 'design', 'layout', 'references', 'mailings', 'review', 'view']
    : [
        'file',
        'home',
        'insert',
        'draw',
        'design',
        'layout',
        'references',
        'mailings',
        'review',
        'view',
      ]
) as readonly string[]
// Word puts an add-in's tab after View; this one shows only where Zotero is around
const ZOTERO_TAB = 'zotero'
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
  tl2br: { label: 'ribbonTableDiagonalDownBorders', Icon: IconBorderDiagonalDown },
  tr2bl: { label: 'ribbonTableDiagonalUpBorders', Icon: IconBorderDiagonalUp },
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
/** Shape Format > Text Direction: what wps:bodyPr vert says, '' for horizontal */
const TEXTBOX_DIRECTION_LABELS: Record<string, StringKey> = {
  '': 'ribbonTextDirectionHorizontal',
  eaVert: 'ribbonDocVertical',
  vert: 'ribbonTextDirectionRotate90',
  vert270: 'ribbonTextDirectionRotate270',
  wordArtVert: 'ribbonTextDirectionStacked',
}
const IMAGE_TABS = ['pictureFormat'] as const
const SHAPE_TABS = ['shapeFormat'] as const
/** Shape Format ▸ Arrange ordering: label + glyph + command, in Word order */
const ASIAN_LAYOUT_ITEMS: Array<[ActionCommandId, StringKey]> = [
  ['format-horizontal-in-vertical', 'ribbonHorizontalInVertical'],
  ['format-combine-characters', 'ribbonCombineCharacters'],
  ['format-two-lines-in-one', 'ribbonTwoLinesInOne'],
  ['format-fit-text', 'ribbonFitText'],
]

/**
 * Shape Format ▸ outline weight gallery, in points — the weights Word offers.
 * The model stores CSS px (1 pt = 12700 EMU = 4/3 px), so a pick converts
 * through `outlineWidthPxOf` and the trigger maps the stored width back.
 */
const SHAPE_OUTLINE_WEIGHTS_PT = [
  '0.25',
  '0.5',
  '0.75',
  '1',
  '1.5',
  '2.25',
  '3',
  '4.5',
  '6',
] as const
const PX_PER_PT = 4 / 3

/** a gallery weight in points as CSS px, rounded like the docx parser (1.5 pt -> 2 px) */
const outlineWidthPxOf = (pt: number): number => Math.round(pt * PX_PER_PT * 100) / 100

/** a stored px width back in points, rounded the same way (2 px -> 1.5 pt) */
const outlineWidthPtOf = (px: number): number => Math.round((px / PX_PER_PT) * 100) / 100

/** the gallery option matching a stored width; null when the box has none or an off-gallery one */
const outlineWeightOptionOf = (px: number | null): string | null => {
  if (px == null) return null
  const pt = outlineWidthPtOf(px)
  return SHAPE_OUTLINE_WEIGHTS_PT.find((weight) => Number(weight) === pt) ?? null
}

/**
 * Trigger text for the weight picker: the matching gallery option (the trigger
 * then reads as that option's label), the box's own off-gallery width, or the
 * control name while the shape still carries Word's default width.
 */
function outlineWeightValueOf(px: number | null, tr: TFunc): string {
  if (px == null) return tr('ribbonShapeOutlineWeight')
  return outlineWeightOptionOf(px) ?? tr('ribbonPtValue', { n: String(outlineWidthPtOf(px)) })
}
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
    styleId: 'PlainGrid',
    name: 'Plain Grid',
    headerFill: null,
    band1Fill: null,
    band2Fill: null,
    borderColor: '808080',
  },
  {
    label: 'ribbonTablePresetBlueHeader',
    styleId: 'BlueHeader',
    name: 'Blue Header',
    headerFill: 'D9EAF7',
    band1Fill: null,
    band2Fill: null,
    borderColor: '5B9BD5',
  },
  {
    label: 'ribbonTablePresetBlueBanded',
    styleId: 'BlueBanded',
    name: 'Blue Banded',
    headerFill: 'BDD7EE',
    band1Fill: 'DDEBF7',
    band2Fill: 'FFFFFF',
    borderColor: '9DC3E6',
  },
  {
    label: 'ribbonTablePresetGrayBanded',
    styleId: 'GrayBanded',
    name: 'Gray Banded',
    headerFill: 'D9E1F2',
    band1Fill: 'E7E6E6',
    band2Fill: 'FFFFFF',
    borderColor: 'A5A5A5',
  },
  {
    label: 'ribbonTablePresetGreenHeader',
    styleId: 'GreenHeader',
    name: 'Green Header',
    headerFill: 'E2F0D9',
    band1Fill: null,
    band2Fill: null,
    borderColor: '70AD47',
  },
] as const satisfies ReadonlyArray<{
  label: StringKey
  styleId: string
  name: string
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
  mailings: 'ribbonTabMailings',
  review: 'ribbonTabReview',
  view: 'ribbonTabView',
  zotero: 'zoteroGroup',
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

/** Word's top-right document mode: Viewing wins over Reviewing (read mode hides editing state) */
export type DocMode = 'editing' | 'reviewing' | 'viewing'
export const docModeOf = (readMode: boolean, trackChanges: boolean): DocMode =>
  readMode ? 'viewing' : trackChanges ? 'reviewing' : 'editing'

const DOC_MODE_KEYS = {
  editing: 'ribbonModeEditing',
  reviewing: 'ribbonModeReviewing',
  viewing: 'ribbonModeViewing',
} as const

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

// the caret stays in the document, as in Word; a scripted refocus would also
// make Blink cancel any spell-check pass a later selection change interrupts
function keepDocumentFocus(e: ReactMouseEvent<HTMLDivElement>) {
  const target = e.target as HTMLElement
  if (target.closest('input, select, textarea, [contenteditable="true"], .gs-dd')) return
  if (target.closest('button')) e.preventDefault()
}

function RibbonInner({
  actionsRef,
  onRibbonCollapsed,
  quickActions,
  trailingActions,
  editor,
  formatState: fs,
  hasDoc,
  documentListPresets,
  recentListPresets,
  onTableDialog,
  tableGridlines,
  styles,
  docDefaults,
  onOpen,
  onSave,
  onSaveAs,
  onNameWithAi,
  section,
  onSection,
  activeSection,
  onPageBorderAll,
  onPaperSizeAll,
  mirrorMargins,
  pageColor,
  watermark,
  watermarkPicture,
  watermarkDocImage,
  themeFonts,
  themeColors,
  styleSetSource,
  savedThemes,
  inkTool,
  onInkTool,
  inkPresets,
  inkActiveId,
  onInkPresets,
  onInkActivate,
  inkCount,
  sources,
  citationStyle,
  zoteroNoteFields,
  showZotero,
  onAddSource,
  zoom,
  darkPage,
  tabRequest,
  header,
  mergeSource,
  mergeRecordIndex,
  onMergeRecordIndex,
  mergeRange,
  onMergeRange,
  footer,
  hfEditing,
  showRuler,
  showNav,
  commentCount,
  openCommentCount,
  resolvedCommentCount,
  showComments,
  onToggleComments,
  commentsMode,
  onCommentsMode,
  editTarget,
  commentAtCaret,
  trackChanges,
  onTrackChanges,
  spellcheck,
  revisionDisplay,
  balloonMode,
  revisionCount,
  revisionFilter,
  revisionAuthors,
  showReviewPane,
  isProtected,
  commentsAllowed,
  trackChangesForced,
  protectActive,
  viewMode,
  readMode,
  onReadMode,
  showGrid,
  splitView,
}: RibbonProps) {
  const { t, lang } = useI18n()
  const collapse = useRibbonCollapse('aidocs.ribbonCollapsed', {
    collapse: t('ribbonCollapse'),
    expand: t('ribbonExpand'),
    // ⌥⌘K / Ctrl+Alt+K: the icon-only middle density (see genoffice#362)
    compact: t('ribbonCompact'),
    expandFull: t('ribbonExpand'),
  })
  // The one-click AI actions need text to work on; grey them out on an empty document
  const docEmpty = !hasDoc || fs.docEmpty
  const [tab, setTab] = useState<RibbonTab>('home')
  const [dropdown, setDropdown] = useState<string | null>(null)
  const [graphicFill, setGraphicFill] = useState<string | null>(null)
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
  // Word's Lock aspect ratio is per picture: a newly selected picture starts locked
  const [lockAspect, setLockAspect] = useState(true)
  useEffect(() => setLockAspect(true), [fs.imageKey])

  useEffect(() => {
    if (showZotero) return
    if (lastRegularTab.current === ZOTERO_TAB) lastRegularTab.current = 'home'
    setTab((current) => (current === ZOTERO_TAB ? 'home' : current))
  }, [showZotero])

  // a request for the Draw tab that has not shown yet: a pen turned on with it stays on
  const drawTabPending = useRef(false)
  useEffect(() => {
    if (!tabRequest) return
    if ((TABS as readonly string[]).includes(tabRequest.tab)) {
      const requested = tabRequest.tab as (typeof TABS)[number]
      drawTabPending.current = requested === 'draw'
      lastRegularTab.current = requested
      setTab(requested)
      setDropdown(tabRequest.dropdown ?? null)
    } else if (
      tabRequest.tab === 'tableDesign' ||
      tabRequest.tab === 'pictureFormat' ||
      tabRequest.tab === 'shapeFormat'
    ) {
      setTab(tabRequest.tab)
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
    escape: true,
    inside: () =>
      Array.from(document.querySelectorAll('[data-rb-panel]')).flatMap((panel) => [
        panel,
        panel.parentElement,
      ]),
  })

  // leaving the Draw tab always drops back to text editing, so the drawing
  // overlay never swallows clicks while its controls are off-screen
  useEffect(() => {
    if (tab === 'draw') drawTabPending.current = false
    else if (inkTool !== 'select' && !drawTabPending.current) onInkTool('select')
  }, [tab, inkTool, onInkTool])

  // a focused textbox sub-editor receives text/paragraph formatting instead
  // of the main editor (Word: ribbon acts on the shape's text while inside it)
  const sub = fs.sub
  const ed = sub ?? editor
  // read-only (Restrict Editing / Read Mode): every edit command is fenced here,
  // button disabled states are only the visual layer on top
  const canEdit = hasDoc && fs.editable
  const aiEntry = useRibbonCommand('toggle-ai')
  const stylesPane = useRibbonCommand('styles-pane')
  const runCommand = useRunCommand()
  const docMode = docModeOf(readMode, trackChanges)
  const applyDocMode = (m: DocMode) => {
    if (m === 'viewing') {
      onReadMode(true)
      return
    }
    if (readMode) onReadMode(false)
    if (m === 'reviewing') {
      if (!trackChanges) onTrackChanges(true)
    } else if (trackChanges && !trackChangesForced) {
      onTrackChanges(false)
    }
  }
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
  const inImage = fs.imageSelected && (!sub || fs.inlinePicture)

  useEffect(() => {
    if (inImage && !wasInImage.current) {
      wasInImage.current = true
      setDropdown(null)
      setTab('pictureFormat')
    } else if (!inImage && wasInImage.current) {
      wasInImage.current = false
      setDropdown(null)
      // a header / footer cell picture hands back to the strip's own tab
      setTab((current) =>
        current === 'pictureFormat' ? (inHf ? 'headerFooter' : lastRegularTab.current) : current,
      )
    }
  }, [inImage, inHf])

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

  /** the selected node's boxes and the one the ribbon addresses (the pressed box of a multi-drawing paragraph) */
  const selectedBoxes = (): { boxes: TextboxDisplay[]; index: number } | null => {
    const boxes = editor.getAttributes('docProtected')?.textboxes as TextboxDisplay[] | null
    if (!Array.isArray(boxes) || boxes.length === 0) return null
    return { boxes, index: selectedBoxIndex(editor.state) }
  }
  const writeSelectedBox = (boxes: TextboxDisplay[], index: number, box: TextboxDisplay) =>
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { textboxes: boxes.map((b, i) => (i === index ? box : b)) })
      .run()

  /**
   * apply fill/outline/weight to the selected floating box. Null
   * deletes the attribute so the file keeps Word's own default instead of an
   * explicit no-op value.
   */
  const setShapeStyle = (patch: {
    fill?: string | null
    borderColor?: string | null
    borderWidthPx?: number | null
  }) => {
    if (!canEdit) return
    const sel = selectedBoxes()
    if (!sel) return
    const box = { ...sel.boxes[sel.index] }
    if ('fill' in patch) {
      if (patch.fill) box.fill = patch.fill
      else delete box.fill
    }
    if ('borderColor' in patch) {
      if (patch.borderColor) box.borderColor = patch.borderColor
      else delete box.borderColor
    }
    if ('borderWidthPx' in patch) {
      if (patch.borderWidthPx) box.borderWidthPx = patch.borderWidthPx
      else delete box.borderWidthPx
    }
    writeSelectedBox(sel.boxes, sel.index, box)
  }

  /** set the selected shape's outline width from one of Word's weights, in points */
  const setShapeOutlineWidth = (pt: number) => {
    setShapeStyle({ borderWidthPx: outlineWidthPxOf(pt) })
  }

  /**
   * Rewrite every run and paragraph of the selected shape. Only for object mode:
   * with the shape selected there is no text selection for a mark command to act
   * on, so Word reformats the whole shape and `editRun`/`editPara` see each part
   * of it in turn. While a sub-editor holds focus the ordinary mark and alignment
   * commands already target the text, and shapeTextCommand routes there instead.
   *
   * Confined to the addressed box like setShapeStyle: a paragraph anchoring
   * several shapes packs them into one docProtected node, and the ribbon reads
   * and writes only the pressed one — reformatting the rest would hit shapes it
   * is not showing.
   *
   * The node view re-feeds its sub-editors from the new attrs, so the shape
   * repaints even while its text is being edited.
   */
  const setShapeText = (
    editRun?: (run: Run) => Run,
    editPara?: (para: TextboxParaDisplay) => TextboxParaDisplay,
  ) => {
    if (!canEdit) return
    const sel = selectedBoxes()
    if (!sel) return
    const current = sel.boxes[sel.index]
    if (current.readOnly) return
    const box = {
      ...current,
      paras: current.paras.map((para) => {
        const withRuns = editRun ? { ...para, runs: para.runs.map(editRun) } : para
        return editPara ? editPara(withRuns) : withRuns
      }),
    }
    writeSelectedBox(sel.boxes, sel.index, box)
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

  /** Set one side of the picture; with the aspect ratio locked the other side follows */
  const setPictureSize = (dim: 'w' | 'h', twips: number) => {
    if (!canEdit || !(twips > 0)) return
    const w = fs.imageWidthPx ?? 0
    const h = fs.imageHeightPx ?? 0
    if (!w || !h) return
    const px = twipsToPx(clampPictureTwips(twips))
    const side = Math.max(1, Math.round(px))
    const next =
      dim === 'w'
        ? { w: side, h: lockAspect ? Math.max(1, Math.round((px * h) / w)) : h }
        : { w: lockAspect ? Math.max(1, Math.round((px * w) / h)) : w, h: side }
    setPicturePx(next.w, next.h)
  }

  /** a cell picture resets to its cell's width at most (Word clips a wider one): the
   *  text column stands in for every other host */
  const inlinePictureHostWidthPx = (): number => {
    if (!fs.inlinePicture) return sectionContentWidthPx
    const sel = ed.state.selection
    const dom =
      sel instanceof NodeSelection ? (ed.view.nodeDOM(sel.from) as HTMLElement | null) : null
    const cell = dom?.closest<HTMLElement>('.page-hf-cell, td, th')
    if (!cell) return sectionContentWidthPx
    const cs = getComputedStyle(cell)
    const inner = cell.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight)
    return inner > 0 ? inner : sectionContentWidthPx
  }

  /** inline run pictures size through their own node; block pictures through docProtected */
  const setPicturePx = (widthPx: number, heightPx: number) => {
    if (fs.inlinePicture) {
      const { state, view } = ed
      const sel = state.selection
      if (!(sel instanceof NodeSelection)) return
      // keep the picture selected: the Size boxes stay for the next edit
      const tr = state.tr.setNodeMarkup(sel.from, undefined, {
        ...sel.node.attrs,
        widthPx,
        heightPx,
      })
      view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, sel.from)))
      view.focus()
      return
    }
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageWidthPx: widthPx, imageHeightPx: heightPx })
      .run()
  }

  /** Word's Reset Size: the picture's own size at the density its file states. A picture wider
   *  than the text column is fitted to it (the page cannot show the overflow). */
  const resetPictureSize = async () => {
    if (!canEdit) return
    const url = fs.imageDataUrl
    if (!inImage || !url) return
    try {
      const natural = await originalPictureSizeOf(url)
      const scale = Math.min(1, inlinePictureHostWidthPx() / natural.width)
      setPicturePx(
        Math.max(1, Math.round(natural.width * scale)),
        Math.max(1, Math.round(natural.height * scale)),
      )
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

  // Word's Home ▸ Borders is the table gallery whenever the selection sits in a
  // table: the lines go on the selected cells, not on their paragraphs. A
  // focused textbox / header / footer sub-editor is never "in a table".
  const HOME_BORDER_CELL_MODE: Record<string, TableBorderMode> = {
    b: 'bottom',
    t: 'top',
    l: 'left',
    r: 'right',
    tblr: 'all',
    none: 'none',
  }
  const applyHomeBorders = (borders: string | null) => {
    if (!sub && inTableOrSelected(editor.state)) {
      applyCellBorders(HOME_BORDER_CELL_MODE[borders ?? 'none'])
      setDropdown(null)
      return
    }
    setParaAttr({ borders })
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

  const currentSize = fs.fontSizePt
  // computed unconditionally (not inside the dropdown render): cheap, and the
  // render-isolation test uses fontFamiliesFor calls as its render probe
  const fontFamilies = fontFamiliesFor(lang)
  const { families: allSystemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // every candidate stays listed (the machine may genuinely lack some), so the
  // system section is the candidates-deduped remainder: no builtin is listed twice
  const systemFontFamilies = systemFamiliesBesidesCandidates(fontFamilies, allSystemFontFamilies)
  // unset align follows the paragraph direction: start is left in LTR, right in RTL
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
  /** a pick from a Home dropdown: run the command, close the dropdown */
  const pick = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(null)
  }
  const setFont = (name: string) => pick('font-family', name)
  const fontMenuItems: Array<{ key: string; name: string; section: 'b' | 's' }> = [
    ...fontFamilies.map((name) => ({ key: `b:${name}`, name, section: 'b' as const })),
    ...systemFontFamilies.map((name) => ({ key: `s:${name}`, name, section: 's' as const })),
  ]
  const fontMenuWindow = useScrollWindow(dropdown === 'fontFamily')
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
  const bodyRef = useRef<HTMLDivElement>(null)
  /** how many style cards Home shows in the ribbon itself, as Word does in a wide window */
  const [inlineStyles, setInlineStyles] = useState(0)
  const inlineStylesShown = useRef({ now: 0, before: -1, width: 0 })
  // the room is measured without the cards already shown, so one more render settles it
  const fitBody = (body: HTMLElement) => {
    fitRibbon(body)
    const shown = inlineStylesShown.current
    let room = styleCardRoom(body)
    // another tab, or the compact ribbon: a count from there says nothing about a tie
    const comparable = room !== null && !body.closest('.ribbon-compact')
    if (!comparable) shown.before = -1
    // off Home the count is kept, so the cards are there again on the way back
    if (room === null || room === shown.now) return
    // a count that comes straight back at the same width is a rounding tie: keep the smaller, once
    if (room === shown.before && body.clientWidth === shown.width) room = Math.min(room, shown.now)
    if (room === shown.now) return
    inlineStylesShown.current = {
      now: room,
      before: comparable ? shown.now : -1,
      width: body.clientWidth,
    }
    setInlineStyles(room)
  }
  const fitBodyRef = useRef(fitBody)
  fitBodyRef.current = fitBody
  // every render: a tab switch or a new language brings captions not yet measured
  useLayoutEffect(() => {
    if (!bodyRef.current) return
    wrapBigCaptions(bodyRef.current)
    fitBody(bodyRef.current)
  })
  useEffect(() => {
    const body = bodyRef.current
    if (!body || typeof ResizeObserver === 'undefined') return
    const watch = new ResizeObserver(() => fitBodyRef.current(body))
    watch.observe(body)
    return () => watch.disconnect()
  }, [])
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
    if (pt !== null) pick('font-size', String(pt))
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

  const applyStyle = (info: StyleInfo) => runCommand('apply-style', info.styleId)

  /** Style cards, shared by the inline gallery and its overflow menu */
  // a document with no styles to show keeps the Styles button
  const styleCards = Math.min(inlineStyles, galleryEntries.length)
  const renderStyleCards = (inMenu: boolean, limit = Infinity) => {
    const apply = (info: StyleInfo) => {
      applyStyle(info)
      if (inMenu) setDropdown(null)
    }
    return (
      <>
        {galleryEntries.slice(0, limit).map((info) => {
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

  const pickListPreset = (levels: CustomNumberingLevel[], recent: ListPresetKind) => {
    const payload: ListPresetPayload = { levels, recent }
    pick('list-preset', JSON.stringify(payload))
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
          onClick={() => pick('list-level', String(i))}
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
            onClick={() => pickListPreset(levels, kind)}
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
    actionsRef.current.toggleRibbon = collapse.toggle
  })
  useEffect(() => onRibbonCollapsed?.(collapse.collapsed), [onRibbonCollapsed, collapse.collapsed])

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

  // Word's copy / paste formatting chords: ⌥⌘C (and ⇧⌘C) / ⌥⌘V on the Mac,
  // Ctrl+Shift+C / Ctrl+Alt+Shift+V elsewhere, where Ctrl+Shift+V is the
  // paste-and-match-style menu accelerator. Paste targets the selection, or
  // the word at the caret.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey)) return
      const one = e.shiftKey !== e.altKey
      const copy = e.code === 'KeyC' && (IS_MAC ? one : e.shiftKey && !e.altKey)
      const paste = e.code === 'KeyV' && e.altKey && e.shiftKey !== IS_MAC
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
    runCommand('change-case', mode)
    setDropdown(null)
  }

  const pasteAs = (mode?: PasteMode | 'picture') => {
    setDropdown(null)
    runCommand('paste-as', mode)
  }

  const clipboard = (action: 'cut' | 'copy') => {
    if (action === 'cut' && !canEdit) return
    document.execCommand(action)
    ed.commands.focus()
  }

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

  const convertItems = (
    [
      ['export-pdf', 'ribbonConvertPdf'],
      ['export-html', 'ribbonConvertHtml'],
      ['export-markdown', 'ribbonConvertMarkdown'],
      ['export-images', 'ribbonConvertImages'],
    ] as const
  ).map(([command, key]) => (
    <button key={command} disabled={!hasDoc} onClick={() => pick(command)}>
      {t(key)}
    </button>
  ))

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
                <button
                  className="qa-name-with-ai"
                  disabled={!hasDoc}
                  onClick={() => {
                    setDropdown(null)
                    onNameWithAi()
                  }}
                >
                  {t('ribbonNameWithAi')}
                </button>
                <RbMenuCommand cmd="print" onRun={() => setDropdown(null)}>
                  {t('appPrintTitle')} <span className="file-menu-key">Ctrl+P</span>
                </RbMenuCommand>
                <div className="file-menu-sep" />
                {convertItems}
              </div>
            )}
          </div>
        )}
        {quickActions}
        {[...TABS, ...(showZotero ? [ZOTERO_TAB] : [])]
          .filter((tabName) => tabName !== 'file')
          .map((tabName) => (
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
        <div className="ribbon-mode-cluster">
          <div className="ribbon-mode-split">
            <button
              type="button"
              className={`ribbon-mode-btn ${showComments ? 'on' : ''}`}
              disabled={!hasDoc}
              aria-pressed={showComments}
              data-tip={t('ribbonShowCommentsTip', { count: commentCount })}
              onClick={onToggleComments}
            >
              <IconComments size={14} />
              <span className="ribbon-mode-text">{t('ribbonShowComments')}</span>
            </button>
            <div className="mode-menu-wrap">
              <button
                type="button"
                className={`ribbon-mode-btn ${dropdown === 'commentsMode' ? 'open' : ''}`}
                disabled={!hasDoc}
                aria-label={t('ribbonShowComments')}
                aria-haspopup="menu"
                aria-expanded={dropdown === 'commentsMode'}
                onClick={() => setDropdown((v) => (v === 'commentsMode' ? null : 'commentsMode'))}
              >
                <span className="ribbon-mode-caret">▾</span>
              </button>
              {dropdown === 'commentsMode' && (
                <div data-rb-panel="" className="file-menu mode-menu" role="menu">
                  {(['contextual', 'list'] as const).map((m) => (
                    <button
                      key={m}
                      role="menuitemradio"
                      aria-checked={commentsMode === m}
                      onClick={() => {
                        setDropdown(null)
                        onCommentsMode(m)
                      }}
                    >
                      <span className="mode-menu-check">{commentsMode === m ? '✓' : ''}</span>
                      {t(m === 'list' ? 'ribbonShowCommentsList' : 'ribbonShowCommentsContextual')}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
          <div className="mode-menu-wrap">
            <button
              type="button"
              className={`ribbon-mode-btn ${dropdown === 'docMode' ? 'open' : ''}`}
              disabled={!hasDoc}
              onClick={() => setDropdown((v) => (v === 'docMode' ? null : 'docMode'))}
            >
              <span className="ribbon-mode-text">{t(DOC_MODE_KEYS[docMode])}</span>
              <span className="ribbon-mode-caret">▾</span>
            </button>
            {dropdown === 'docMode' && (
              <div data-rb-panel="" className="file-menu mode-menu">
                {(['editing', 'reviewing', 'viewing'] as const).map((m) => (
                  <button
                    key={m}
                    // Editing needs the recorder off (impossible under forced track changes);
                    // Reviewing records changes, so it needs an editable document
                    disabled={
                      m === 'editing'
                        ? !fs.editable || trackChangesForced
                        : m === 'reviewing'
                          ? !fs.editable
                          : false
                    }
                    onClick={() => {
                      setDropdown(null)
                      applyDocMode(m)
                    }}
                  >
                    <span className="mode-menu-check">{docMode === m ? '✓' : ''}</span>
                    {t(DOC_MODE_KEYS[m])}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        {trailingActions}
      </div>

      <div className="ribbon-body" data-ribbon-body="" ref={bodyRef}>
        {tab === 'headerFooter' && hfEditing ? (
          <HeaderFooterTab
            info={hfEditing}
            dropdown={dropdown}
            setDropdown={setDropdown}
            onAction={(action) => runCommand('header-footer', JSON.stringify(action))}
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
                <Dropdown
                  className="rb-wrap-dd"
                  disabled={!canEdit}
                  tip={t('ribbonShapeOutlineWeight')}
                  value={outlineWeightValueOf(fs.shapeBorderWidthPx, t)}
                  options={SHAPE_OUTLINE_WEIGHTS_PT.map((pt) => ({
                    value: pt,
                    label: t('ribbonPtValue', { n: pt }),
                  }))}
                  onPick={(v) => {
                    if (!canEdit) return
                    setShapeOutlineWidth(Number(v))
                  }}
                />
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
                    <div className="rb-row">
                      <Dropdown
                        className="rb-wrap-dd"
                        disabled={!canEdit}
                        tip={t('ribbonTextDirection')}
                        value={
                          fs.shapeVert && !(fs.shapeVert in TEXTBOX_DIRECTION_LABELS)
                            ? 'vert'
                            : (fs.shapeVert ?? '')
                        }
                        options={TEXTBOX_DIRECTIONS.map((vert) => ({
                          value: vert ?? '',
                          label: t(TEXTBOX_DIRECTION_LABELS[vert ?? '']),
                        }))}
                        onPick={(v) => runCommand('shape-text-direction', v || HORIZONTAL)}
                      />
                    </div>
                  </div>
                </div>
                <div className="ribbon-group-label">{t('ribbonGroupText')}</div>
              </div>
            )}
            <div className="ribbon-sep" />
            <ArrangeGroup
              editor={editor}
              hasDoc={canEdit}
              dropdown={dropdown}
              setDropdown={setDropdown}
            />
          </div>
        ) : tab === 'pictureFormat' && inImage && fs.pictureBox ? (
          <div className="table-ribbon-body">
            <ArrangeGroup
              editor={editor}
              hasDoc={canEdit}
              dropdown={dropdown}
              setDropdown={setDropdown}
            />
          </div>
        ) : tab === 'pictureFormat' && inImage ? (
          <div className="table-ribbon-body">
            {fs.graphicSelected && (
              <>
                <div className="ribbon-group">
                  <div className="ribbon-group-items">
                    <div className="rb-split-wrap">
                      <button
                        className="rb-big"
                        disabled={!canEdit}
                        data-tip={t('ribbonGraphicsFill')}
                        // command-exempt: opens this control's own colour palette
                        onClick={() => {
                          if (dropdown === 'graphicsFill') return setDropdown(null)
                          // the fill is read from the picture's SVG, which may have to be fetched
                          void selectedGraphicFill(editor).then((fill) => {
                            setGraphicFill(fill ?? null)
                            setDropdown('graphicsFill')
                          })
                        }}
                      >
                        <span className="rb-big-icon">
                          <IconShading />
                        </span>
                        <span>{t('ribbonGraphicsFill')}</span>
                      </button>
                      {dropdown === 'graphicsFill' && (
                        <RibbonColorPalette
                          current={graphicFill}
                          noneLabel={t('ribbonNoFill')}
                          onPick={(hex) => {
                            runCommand('picture-graphics-fill', hex ?? NONE)
                            setDropdown(null)
                          }}
                        />
                      )}
                    </div>
                  </div>
                  <div className="ribbon-group-label">{t('ribbonGroupGraphicsStyles')}</div>
                </div>
                <div className="ribbon-sep" />
              </>
            )}
            {!fs.inlinePicture && (
              <>
                {/* ---- Adjust: remove background / crop / replace picture ---- */}
                <div className="ribbon-group">
                  <div className="ribbon-group-items">
                    <button
                      className="rb-big"
                      disabled={!canEdit}
                      data-tip={t('ribbonRemoveBgTip')}
                      onClick={() => runCommand('picture-edit', 'cutout')}
                    >
                      <span className="rb-big-icon">
                        <IconRemoveBg size={28} />
                      </span>
                      <span>{t('ribbonRemoveBg')}</span>
                    </button>
                    <button
                      className="rb-big"
                      disabled={!canEdit}
                      data-tip={t('ribbonReplacePictureTip')}
                      onClick={() => runCommand('picture-replace')}
                    >
                      <span className="rb-big-icon">
                        <IconReplacePicture size={28} />
                      </span>
                      <span>{t('ribbonReplacePicture')}</span>
                    </button>
                  </div>
                  <div className="ribbon-group-label">{t('ribbonGroupAdjust')}</div>
                </div>
                <div className="ribbon-sep" />
                {/* ---- the block picture's own left / centre / right, which Word does not have ---- */}
                <div className="ribbon-group">
                  <div className="ribbon-group-items">
                    <div className="table-tool-inset">
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
                            // command-exempt: moved with its group, not new; still sets the picture's own attribute
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
                    </div>
                  </div>
                  <div className="ribbon-group-label">{t('ribbonGroupAlignment')}</div>
                </div>
                <div className="ribbon-sep" />
              </>
            )}
            <ArrangeGroup
              editor={editor}
              hasDoc={canEdit}
              dropdown={dropdown}
              setDropdown={setDropdown}
            />
            <div className="ribbon-sep" />
            {/* ---- Size: height/width in the measurement unit, aspect lock, reset ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                {!fs.inlinePicture && (
                  <button
                    className="rb-big"
                    disabled={!canEdit}
                    data-tip={t('ribbonCropTip')}
                    onClick={() => runCommand('picture-edit', 'crop')}
                  >
                    <span className="rb-big-icon">
                      <IconCrop size={28} />
                    </span>
                    <span>{t('ribbonCrop')}</span>
                  </button>
                )}
                <div className="table-tool-inset">
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
                    <button
                      data-tip={t('ribbonResetSizeTip')}
                      // command-exempt: moved with its group, not new; reads the picture file for its size
                      onClick={() => void resetPictureSize()}
                    >
                      {t('ribbonResetSize')}
                    </button>
                  </div>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupSize')}</div>
            </div>
          </div>
        ) : tab === 'tableDesign' ? (
          <div className="table-ribbon-body">
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
                    onClick={() => runCommand('table-look', `${key}=${!tableLook[key]}`)}
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
              <div className="table-style-gallery">
                <button
                  className="table-style-card"
                  data-tip={t('ribbonRemoveTableStyleTip')}
                  onClick={() => runCommand('table-style', NONE)}
                >
                  <span className="table-style-card-grid plain" />
                  <span>{t('ribbonNoStyle')}</span>
                </button>
                {TABLE_PRESETS.map((preset) => (
                  <button
                    key={preset.label}
                    className={
                      tableAttrs.tblStyleId === preset.styleId
                        ? 'table-style-card active'
                        : 'table-style-card'
                    }
                    data-tip={t(preset.label)}
                    onClick={() => runCommand('table-preset', JSON.stringify(preset))}
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
                  .filter(
                    (info) =>
                      info.type === 'table' &&
                      info.styleId !== 'TableNormal' &&
                      !TABLE_PRESETS.some((preset) => preset.styleId === info.styleId),
                  )
                  .map((info) => (
                    <button
                      key={info.styleId}
                      className={
                        tableAttrs.tblStyleId === info.styleId
                          ? 'table-style-card active'
                          : 'table-style-card'
                      }
                      data-tip={t('ribbonApplyTableStyleTip', { name: info.name })}
                      onClick={() => runCommand('table-style', info.styleId)}
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
              <div className="rb-split-wrap table-split">
                <button
                  className="table-split-main"
                  data-tip={t(TABLE_BORDER_ITEMS[lastBorderMode].label)}
                  // command-exempt: moved with its group, not new; repeats the last border choice
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
              <div className="ribbon-group-label">{t('ribbonGroupBorders')}</div>
            </div>
          </div>
        ) : tab === 'tableLayout' ? (
          <div className="table-ribbon-body">
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <div className="rb-split-wrap">
                  <button
                    className={`rb-big${dropdown === 'tableSelect' ? ' active' : ''}`}
                    aria-expanded={dropdown === 'tableSelect'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableSelect' ? null : 'tableSelect'))
                    }
                  >
                    <span className="rb-big-icon">
                      <IconSelectCells size={28} />
                      <IconCaret />
                    </span>
                    <span>{t('ribbonSelect')}</span>
                  </button>
                  {dropdown === 'tableSelect' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      {TABLE_SELECT_ITEMS.map(([kind, label]) => (
                        <button key={kind} onClick={() => pick(`table-select-${kind}`)}>
                          <IconSelectCells size={17} />
                          <span>{t(label)}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <button
                  className={tableGridlines ? 'rb-big active' : 'rb-big'}
                  aria-pressed={!!tableGridlines}
                  data-tip={t('ribbonViewGridlinesTip')}
                  onClick={() => runCommand('table-gridlines')}
                >
                  <span className="rb-big-icon">
                    <IconGridlines size={28} />
                  </span>
                  <span>{t('ribbonViewGridlines')}</span>
                </button>
                <button className="rb-big" onClick={() => pick('table-properties')}>
                  <span className="rb-big-icon">
                    <IconTableProperties size={28} />
                  </span>
                  <span>{t('ribbonTableProperties')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonTableData')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <div className="rb-split-wrap">
                  <button
                    className={`rb-big${dropdown === 'tableDelete' ? ' active' : ''}`}
                    aria-expanded={dropdown === 'tableDelete'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableDelete' ? null : 'tableDelete'))
                    }
                  >
                    <span className="rb-big-icon">
                      <IconTableDelete size={28} />
                      <IconCaret />
                    </span>
                    <span>{t('ribbonTableDeleteMenu')}</span>
                  </button>
                  {dropdown === 'tableDelete' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      <button onClick={() => pick('table-delete-cells')}>
                        <IconDeleteCells size={17} />
                        <span>{t('ribbonDeleteCells')}</span>
                      </button>
                      {(
                        [
                          ['ribbonDeleteColumn', IconColDelete, 'table-delete-columns'],
                          ['ribbonDeleteRow', IconRowDelete, 'table-delete-rows'],
                          ['ribbonDeleteTable', IconTableDelete, 'table-delete-table'],
                        ] as Array<
                          [StringKey, (props: { size?: number }) => ReactNode, ActionCommandId]
                        >
                      ).map(([label, Icon, command]) => (
                        <button key={label} onClick={() => pick(command)}>
                          <Icon size={17} />
                          <span>{t(label)}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <button className="rb-big" onClick={() => runCommand('table-insert-rows-above')}>
                  <span className="rb-big-icon">
                    <IconRowInsertAbove size={28} />
                  </span>
                  <span>{t('ribbonInsertAbove')}</span>
                </button>
                <button className="rb-big" onClick={() => runCommand('table-insert-rows-below')}>
                  <span className="rb-big-icon">
                    <IconRowInsertBelow size={28} />
                  </span>
                  <span>{t('ribbonInsertBelow')}</span>
                </button>
                <button className="rb-big" onClick={() => runCommand('table-insert-cols-left')}>
                  <span className="rb-big-icon">
                    <IconColInsertLeft size={28} />
                  </span>
                  <span>{t('ribbonInsertLeft')}</span>
                </button>
                <button className="rb-big" onClick={() => runCommand('table-insert-cols-right')}>
                  <span className="rb-big-icon">
                    <IconColInsertRight size={28} />
                  </span>
                  <span>{t('ribbonInsertRight')}</span>
                </button>
                <button className="rb-big" onClick={() => runCommand('table-insert-cells')}>
                  <span className="rb-big-icon">
                    <IconInsertCells size={28} />
                  </span>
                  <span>{t('ribbonInsertCells')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupRowsCols')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <button
                  className="rb-big"
                  disabled={!fs.canMergeCells}
                  onClick={() => runCommand('table-merge-cells')}
                >
                  <span className="rb-big-icon">
                    <IconMergeCells size={28} />
                  </span>
                  <span>{t('ribbonMergeCells')}</span>
                </button>
                <button className="rb-big" onClick={() => runCommand('table-split-cells')}>
                  <span className="rb-big-icon">
                    <IconSplitCells size={28} />
                  </span>
                  <span>{t('ribbonSplitCells')}</span>
                </button>
                <button
                  className="rb-big"
                  data-tip={t('ribbonSplitTableTip')}
                  disabled={!fs.canSplitTable}
                  onClick={() => runCommand('table-split-table')}
                >
                  <span className="rb-big-icon">
                    <IconSplitTable size={28} />
                  </span>
                  <span>{t('ribbonSplitTable')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupMerge')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <div className="rb-split-wrap">
                  <button
                    className={`rb-big${dropdown === 'tableAutoFit' ? ' active' : ''}`}
                    data-tip={`${t('ribbonAutoFit')}: ${t(tableAutoFitLabel)}`}
                    aria-expanded={dropdown === 'tableAutoFit'}
                    onClick={() =>
                      setDropdown((current) => (current === 'tableAutoFit' ? null : 'tableAutoFit'))
                    }
                  >
                    <span className="rb-big-icon">
                      <IconAutoFit size={28} />
                      <IconCaret />
                    </span>
                    <span>{t('ribbonAutoFit')}</span>
                  </button>
                  {dropdown === 'tableAutoFit' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      {TABLE_AUTO_FIT_OPTIONS.map(([mode, label]) => (
                        <button
                          key={mode}
                          className={tableAutoFitMode === mode ? 'active' : ''}
                          onClick={() => pick(`table-autofit-${mode}`)}
                        >
                          <IconAutoFit size={17} />
                          <span>{t(label)}</span>
                          <span className="table-menu-state" aria-hidden="true" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <div className="table-tool-inset">
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
                </div>
                <div className="rb-col rb-pair">
                  <button
                    className="rb-small"
                    data-tip={t('ribbonDistributeRows')}
                    onClick={() => runCommand('table-distribute-rows')}
                  >
                    <IconDistributeRows size={14} />
                    <span className="rb-small-label">{t('ribbonDistributeRows')}</span>
                  </button>
                  <button
                    className="rb-small"
                    data-tip={t('ribbonDistributeColumns')}
                    onClick={() => runCommand('table-distribute-columns')}
                  >
                    <IconDistributeColumns size={14} />
                    <span className="rb-small-label">{t('ribbonDistributeColumns')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupCellSize')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <div className="table-tool-inset">
                  <div
                    className="cell-align-grid"
                    role="group"
                    aria-label={t('ribbonCellAlignment')}
                  >
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
                        onClick={() => runCommand('table-cell-align', `${v},${h}`)}
                      >
                        <IconCellAlign v={v} h={h} size={16} />
                      </button>
                    ))}
                  </div>
                </div>
                <div className="rb-split-wrap">
                  <button
                    className={`rb-big${dropdown === 'cellTextDirection' ? ' active' : ''}`}
                    data-tip={t('ribbonTextDirection')}
                    aria-expanded={dropdown === 'cellTextDirection'}
                    onClick={() =>
                      setDropdown((current) =>
                        current === 'cellTextDirection' ? null : 'cellTextDirection',
                      )
                    }
                  >
                    <span className="rb-big-icon">
                      <IconTextDirection size={28} />
                      <IconCaret />
                    </span>
                    <span>{t('ribbonTextDirection')}</span>
                  </button>
                  {dropdown === 'cellTextDirection' && (
                    <div data-rb-panel="" className="layout-menu table-autofit-menu">
                      {CELL_TEXT_DIRECTIONS.map(([dir, labelKey]) => (
                        <button
                          key={dir}
                          className={(fs.cellTextDirection ?? 'lrTb') === dir ? 'active' : ''}
                          onClick={() => pick('table-text-direction', dir)}
                        >
                          <IconTextDirection size={17} />
                          <span>{t(labelKey)}</span>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                <button
                  className="rb-big"
                  // command-exempt: moved with its group, not new; opens the table's own dialog
                  onClick={() => {
                    setDropdown(null)
                    if (canEdit) onTableDialog?.('cellMargins')
                  }}
                >
                  <span className="rb-big-icon">
                    <IconCellMargins size={28} />
                  </span>
                  <span>{t('ribbonCellMarginsCmd')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupAlignment')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <RbCommand cmd="sort" look="big">
                  <IconSort size={28} />
                </RbCommand>
                <button
                  className={tableHeader.active ? 'rb-big active' : 'rb-big'}
                  disabled={!tableHeader.enabled}
                  aria-pressed={tableHeader.active}
                  onClick={() => runCommand('table-repeat-header')}
                >
                  <span className="rb-big-icon">
                    <IconRepeatHeader size={28} />
                  </span>
                  <span>{t('ribbonRepeatHeaderRows')}</span>
                </button>
                <RbCommand cmd="table-convert-to-text" look="big">
                  <IconConvert size={28} />
                </RbCommand>
                <RbCommand cmd="table-formula" look="big">
                  <span className="rb-big-glyph">fx</span>
                </RbCommand>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupData')}</div>
            </div>
          </div>
        ) : tab === 'home' ? (
          <>
            {/* ---- Genspark AI (first slot: entry + one-click AI actions) ---- */}
            <div className="ribbon-group" data-tight="1000">
              <div className="ribbon-group-items">
                <button
                  className={`rb-big ai-entry ${aiEntry.pressed ? 'active' : ''}`}
                  data-tip={t(aiEntry.control.tip)}
                  onClick={aiEntry.run}
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
                  onClick={() => runCommand('ai-preset', t('aiSummarizePrompt'))}
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
                    runCommand(
                      'ai-preset',
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
                  onClick={() => runCommand('ai-preset', t('aiTidyPrompt'))}
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
                <button
                  className="rb-big ai-entry"
                  disabled={docEmpty}
                  data-tip={t('aiMindMapBtn')}
                  onClick={() => runCommand('ai-preset', t('aiMindMapPrompt'))}
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
                        <rect x="2.5" y="9.5" width="6" height="5" rx="1.5" />
                        <rect x="15.5" y="4" width="6" height="5" rx="1.5" />
                        <rect x="15.5" y="15" width="6" height="5" rx="1.5" />
                        <path d="M8.5 12h3.5M12 12V6.5h3.5M12 12v5.5h3.5" />
                      </svg>
                    </span>
                  </span>
                  <span>{t('aiMindMapBtn')}</span>
                </button>
              </div>
              <div className="ribbon-group-label">Genspark AI</div>
            </div>

            <div className="ribbon-sep" />

            <div className="ribbon-group" data-tight="1000">
              <div className="ribbon-group-items">
                <div className="rb-split-wrap">
                  <button
                    className={`rb-big${dropdown === 'convert' ? ' active' : ''}`}
                    disabled={!hasDoc}
                    data-tip={t('ribbonConvertTip')}
                    aria-expanded={dropdown === 'convert'}
                    onClick={() => setDropdown((v) => (v === 'convert' ? null : 'convert'))}
                  >
                    <span className="rb-big-icon">
                      <IconConvert size={22} />
                      <IconCaret />
                    </span>
                    <span>{t('ribbonConvert')}</span>
                  </button>
                  {dropdown === 'convert' && (
                    <div data-rb-panel="" className="spacing-menu paste-menu" role="menu">
                      {convertItems}
                    </div>
                  )}
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonConvert')}</div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Clipboard ---- */}
            <div className="ribbon-group" data-tight="1001">
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
                      <button
                        role="menuitem"
                        onClick={() => {
                          setDropdown(null)
                          runCommand('paste-as', 'picture')
                        }}
                      >
                        <IconPicture size={16} />
                        {t('ribbonPasteFormatPicture')}
                      </button>
                      <div className="spacing-menu-sep" />
                      <button role="menuitem" onClick={() => pick('edit-paste-special')}>
                        {t('ribbonPasteSpecial')}
                      </button>
                      <button role="menuitem" onClick={() => pick('autocorrect-options')}>
                        {t('ribbonSetDefaultPaste')}
                      </button>
                    </div>
                  )}
                </div>
                <div className="rb-col">
                  <button
                    className="rb-small"
                    disabled={clipboardOff('cut', { hasDoc, target: editTarget })}
                    data-tip={t('ribbonCutTip')}
                    aria-label={t('ribbonCutTip')}
                    onClick={() => clipboard('cut')}
                  >
                    <IconCut />
                  </button>
                  <button
                    className="rb-small"
                    disabled={clipboardOff('copy', { hasDoc, target: editTarget })}
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
            <div className="ribbon-group" data-menu={HOME_FOLDS ? 1004 : undefined}>
              {HOME_FOLDS && (
                <RbGroupMenu name={t('appFontDialogTitle')}>
                  <IconFontColorA size={BIG} />
                </RbGroupMenu>
              )}
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
                      <div
                        data-rb-panel=""
                        className="spacing-menu rb-font-family-menu"
                        ref={fontMenuWindow.ref}
                        onScroll={fontMenuWindow.onScroll}
                      >
                        {(['b', 's'] as const).map((section) => {
                          const items = fontMenuItems.filter((it) => it.section === section)
                          if (items.length === 0) return null
                          const offset = section === 'b' ? 0 : fontFamilies.length
                          return (
                            <div key={section}>
                              <div className="rb-menu-group-label">
                                {t(section === 'b' ? 'ribbonFontsCommon' : 'ribbonFontsSystem')}
                              </div>
                              {items.map((it, i) => (
                                <button
                                  key={it.key}
                                  data-font-key={it.key}
                                  className={`${it.name === currentFont ? 'active' : ''}${
                                    it.key === fontHighlight ? ' kbd-focus' : ''
                                  }`}
                                  // symbol fonts would render their own name as pictographs
                                  style={{
                                    fontFamily:
                                      isSymbolFontFamily(it.name) ||
                                      !fontMenuWindow.inWindow(offset + i)
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
                            onClick={() => pick('font-size', String(o.pt))}
                          >
                            {o.name}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  <RbCommand cmd="font-grow">
                    <IconGrowFont />
                  </RbCommand>
                  <RbCommand cmd="font-shrink">
                    <IconShrinkFont />
                  </RbCommand>
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
                  <RbCommand cmd="clear-formatting">
                    <IconClearFormat />
                  </RbCommand>
                  <RbCommand cmd="format-phonetic-guide">
                    <IconPhoneticGuide />
                  </RbCommand>
                  <RbCommand cmd="character-border">
                    <IconCharBorder />
                  </RbCommand>
                </div>
                <div className="rb-row">
                  <RbCommand cmd="bold">
                    <b>B</b>
                  </RbCommand>
                  <RbCommand cmd="italic">
                    <i>I</i>
                  </RbCommand>
                  <RbCommand cmd="underline">
                    <u>U</u>
                  </RbCommand>
                  <RbCommand cmd="strikethrough">
                    <s>ab</s>
                  </RbCommand>
                  <RbCommand cmd="subscript">
                    <IconSubscript />
                  </RbCommand>
                  <RbCommand cmd="superscript">
                    <IconSuperscript />
                  </RbCommand>
                  <RbTextEffects
                    open={dropdown === 'textFx'}
                    onToggle={() => setDropdown((v) => (v === 'textFx' ? null : 'textFx'))}
                    onClose={() => setDropdown(null)}
                  />
                  <span className="rb-mini-sep" />
                  {/* highlight: main button applies pen color, caret opens palette */}
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon rb-color-btn ${fs.highlight ? 'active' : ''}`}
                      disabled={!canEdit}
                      data-tip={t('ribbonTextHighlightColor')}
                      aria-label={t('ribbonTextHighlightColor')}
                      onClick={() =>
                        pick('highlight', fs.highlight === penHighlight ? NONE : penHighlight)
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
                                pick('highlight', h)
                              }}
                            />
                          ))}
                        </div>
                        <button
                          className={`color-none color-highlight-none ${!fs.highlight ? 'selected' : ''}`}
                          onClick={() => pick('highlight', NONE)}
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
                      onClick={() => pick('font-color', penColor ?? NONE)}
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
                          setPenColor(hex)
                          pick('font-color', hex ?? NONE)
                        }}
                      />
                    )}
                  </div>
                  <RbCommand cmd="character-shading">
                    <IconCharShading />
                  </RbCommand>
                  <RbCommand cmd="format-enclose-characters">
                    <IconEncloseChars />
                  </RbCommand>
                </div>
              </div>
            </div>

            <div className="ribbon-sep" />

            {/* ---- Paragraph ---- */}
            <div className="ribbon-group" data-menu={HOME_FOLDS ? 1003 : undefined}>
              {HOME_FOLDS && (
                <RbGroupMenu name={t('ribbonGroupParagraph')}>
                  <IconAlignLeft size={BIG} />
                </RbGroupMenu>
              )}
              <div className="ribbon-group-items rb-font-group">
                <div className="rb-row">
                  <div className="rb-split-wrap">
                    <RbCommand cmd="bullets">
                      <IconBullets />
                    </RbCommand>
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
                          onClick={() => pick('list-none')}
                        >
                          {t('ribbonListNone')}
                        </button>
                        {listLibrary('bullets').map((levels, i) => (
                          <button
                            key={i}
                            className="list-gallery-card list-gallery-glyph"
                            onClick={() => pickListPreset(levels, 'bullets')}
                          >
                            {bulletPreview(levels)}
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
                          onClick={() => pick('list-dialog', 'bullet')}
                        >
                          {t('ribbonDefineNewBullet')}…
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <RbCommand cmd="numbering">
                      <IconNumbered />
                    </RbCommand>
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
                          onClick={() => pick('list-none')}
                        >
                          {t('ribbonListNone')}
                        </button>
                        {listLibrary('numbers').map((levels, i) => (
                          <button
                            key={i}
                            className="list-gallery-card list-gallery-preview"
                            onClick={() => pickListPreset(levels, 'numbers')}
                          >
                            {numberPreview(levels)}
                          </button>
                        ))}
                        {presetSection(
                          'ribbonDocumentNumbers',
                          documentListPresets?.numbers,
                          'numbers',
                          numberPreview,
                        )}
                        {listLevelRow}
                        <button
                          className="list-gallery-define"
                          onClick={() => pick('list-dialog', 'number')}
                        >
                          {t('ribbonDefineNewNumber')}…
                        </button>
                        <button
                          className="list-gallery-define"
                          disabled={!fs.listOrdered}
                          onClick={() => pick('list-dialog', 'value')}
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
                        {listLibrary('multi').map((levels, i) => (
                          <button
                            key={i}
                            className="list-gallery-card list-gallery-card-multi"
                            onClick={() => pickListPreset(levels, 'multi')}
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
                          onClick={() => pick('list-dialog', 'multi')}
                        >
                          {t('ribbonDefineNewList')}…
                        </button>
                      </div>
                    )}
                  </div>
                  <span className="rb-mini-sep" />
                  <RbCommand cmd="indent-decrease">
                    <IconIndentDec />
                  </RbCommand>
                  <RbCommand cmd="indent-increase">
                    <IconIndentInc />
                  </RbCommand>
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon"
                      disabled={!canEdit}
                      data-tip={t('ribbonAsianLayout')}
                      aria-label={t('ribbonAsianLayout')}
                      onClick={() => setDropdown(dropdown === 'asianlayout' ? null : 'asianlayout')}
                    >
                      <IconAsianLayout />
                    </button>
                    {dropdown === 'asianlayout' && (
                      <div data-rb-panel="" className="layout-menu">
                        {ASIAN_LAYOUT_ITEMS.map(([id, label]) => (
                          <RbMenuCommand key={id} cmd={id} onRun={() => setDropdown(null)}>
                            {t(label)}
                          </RbMenuCommand>
                        ))}
                      </div>
                    )}
                  </div>
                  <RbCommand cmd="sort">
                    <IconSort />
                  </RbCommand>
                  <span className="rb-mini-sep" />
                  <RbCommand cmd="show-marks">
                    <IconPilcrow />
                  </RbCommand>
                </div>
                <div className="rb-row">
                  <RbCommand cmd="align-left">
                    <IconAlignLeft />
                  </RbCommand>
                  <RbCommand cmd="align-center">
                    <IconAlignCenter />
                  </RbCommand>
                  <RbCommand cmd="align-right">
                    <IconAlignRight />
                  </RbCommand>
                  <RbCommand cmd="align-justify">
                    <IconAlignJustify />
                  </RbCommand>
                  <RbCommand cmd="align-distributed">
                    <IconAlignDistributed />
                  </RbCommand>
                  {showDirection && (
                    <>
                      <span className="rb-mini-sep" />
                      <RbCommand cmd="direction-ltr">
                        <IconDirLtr />
                      </RbCommand>
                      <RbCommand cmd="direction-rtl">
                        <IconDirRtl />
                      </RbCommand>
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
                            onClick={() => pick('line-spacing', String(s))}
                          >
                            {s.toFixed(2).replace(/0+$/, '').replace(/\.$/, '.0')}
                          </button>
                        ))}
                        <div className="spacing-menu-sep" />
                        <button onClick={() => pick('paragraph-dialog')}>
                          {t('ribbonLineSpacingOptions')}
                        </button>
                        <button onClick={() => pick('space-before')}>
                          {t(
                            fs.spaceBefore > 0 ? 'ribbonRemoveSpaceBefore' : 'ribbonAddSpaceBefore',
                          )}
                        </button>
                        <button onClick={() => pick('space-after')}>
                          {t(fs.spaceAfter > 0 ? 'ribbonRemoveSpaceAfter' : 'ribbonAddSpaceAfter')}
                        </button>
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
                        onPick={(hex) => pick('paragraph-shading', hex ?? NONE)}
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
                        <button onClick={() => applyHomeBorders('b')}>
                          {t('ribbonBorderBottom')}
                        </button>
                        <button onClick={() => applyHomeBorders('t')}>
                          {t('ribbonBorderTop')}
                        </button>
                        <button onClick={() => applyHomeBorders('l')}>
                          {t('ribbonBorderLeft')}
                        </button>
                        <button onClick={() => applyHomeBorders('r')}>
                          {t('ribbonBorderRight')}
                        </button>
                        <button onClick={() => applyHomeBorders(null)}>
                          {t('ribbonNoBorders')}
                        </button>
                        <button
                          // command-exempt: borders the paragraph or the selected cells, which no one command covers
                          onClick={() => applyHomeBorders('tblr')}
                        >
                          {t('ribbonBorderBox')}
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
            <div className="ribbon-group ribbon-group-styles" data-tight="1001">
              <div className="ribbon-group-items rb-split-wrap style-gallery-wrap">
                {styleCards > 0 && (
                  <div className="style-gallery-inline">{renderStyleCards(false, styleCards)}</div>
                )}
                <button
                  className={`${styleCards > 0 ? 'rb-styles-more' : 'rb-big'} rb-styles-gallery${dropdown === 'styleGallery' ? ' active' : ''}`}
                  disabled={!canEdit}
                  aria-haspopup="menu"
                  aria-expanded={dropdown === 'styleGallery'}
                  aria-label={t('ribbonGroupStyles')}
                  data-tip={t('ribbonStylesTip')}
                  onClick={() => setDropdown((v) => (v === 'styleGallery' ? null : 'styleGallery'))}
                >
                  {styleCards > 0 ? (
                    <IconCaret />
                  ) : (
                    <>
                      <span className="rb-big-icon rb-styles-gallery-icon">
                        <IconStyles size={22} />
                        <span className="rb-caret-inline">
                          <IconCaret />
                        </span>
                      </span>
                      <span>{t('ribbonGroupStyles')}</span>
                    </>
                  )}
                </button>
                {dropdown === 'styleGallery' && (
                  <div data-rb-panel="" className="style-gallery-menu">
                    <div className="style-gallery-grid">{renderStyleCards(true)}</div>
                    <button
                      className="style-gallery-clear"
                      onClick={() => pick('clear-formatting')}
                    >
                      <IconClearFormat />
                      {t('ribbonClearFormatting')}
                    </button>
                    <button
                      className="style-gallery-clear style-gallery-pane"
                      disabled={stylesPane.disabled}
                      onClick={() => pick('styles-pane')}
                    >
                      <IconStylesPane size={16} />
                      {t('ribbonStylesPane')}
                    </button>
                  </div>
                )}
                <button
                  className={`rb-big rb-styles-pane${stylesPane.pressed ? ' active' : ''}`}
                  data-drop={HOME_FOLDS ? 1002 : undefined}
                  disabled={stylesPane.disabled}
                  aria-pressed={stylesPane.pressed}
                  data-tip={t(stylesPane.control.tip)}
                  onClick={stylesPane.run}
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
                    <RbCommand cmd="find" look="small">
                      <IconSearch />
                    </RbCommand>
                    <RbCommand cmd="replace" look="small">
                      <IconReplace />
                    </RbCommand>
                    <button
                      className="rb-small"
                      disabled={!hasDoc}
                      data-tip={t('ribbonSelectAllTip')}
                      onClick={() => ed.chain().focus().selectAll().run()}
                    >
                      <IconSelectAll />
                      <span className="rb-small-label">{t('ribbonSelectAll')}</span>
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
            dropdown={dropdown}
            setDropdown={setDropdown}
            tool={inkTool}
            onTool={(tool) => runCommand('ink-tool', tool)}
            presets={inkPresets}
            activeId={inkActiveId}
            onPresets={onInkPresets}
            onActivate={onInkActivate}
            annotationCount={inkCount}
            onClearAll={() => runCommand('ink-clear')}
          />
        ) : tab === 'insert' ? (
          <InsertTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            header={header}
            footer={footer}
            isProtected={isProtected}
            commentsAllowed={commentsAllowed}
          />
        ) : tab === 'design' ? (
          <DesignTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            pageColor={pageColor}
            section={section}
            onSection={onSection}
            onPageBorderAll={onPageBorderAll}
            watermark={watermark}
            watermarkPicture={watermarkPicture}
            watermarkDocImage={watermarkDocImage}
            themeFonts={themeFonts}
            themeColors={themeColors}
            styleSetSource={styleSetSource ?? null}
            savedThemes={savedThemes ?? NO_SAVED_THEMES}
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
            onPaperSizeAll={onPaperSizeAll}
            mirrorMargins={mirrorMargins}
          />
        ) : tab === 'references' ? (
          <ReferencesTab
            editor={editor}
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            sources={sources}
            citationStyle={citationStyle ?? null}
            onAddSource={onAddSource}
          />
        ) : tab === ZOTERO_TAB ? (
          <ZoteroTab
            hasDoc={canEdit}
            setDropdown={setDropdown}
            zoteroNoteFields={zoteroNoteFields}
          />
        ) : tab === 'mailings' ? (
          <MailingsTab
            hasDoc={canEdit}
            dropdown={dropdown}
            setDropdown={setDropdown}
            source={mergeSource}
            recordIndex={mergeRecordIndex}
            onRecordIndex={onMergeRecordIndex}
            range={mergeRange}
            onRange={onMergeRange}
          />
        ) : tab === 'review' ? (
          <ReviewTab
            editor={editor}
            hasDoc={hasDoc}
            dropdown={dropdown}
            setDropdown={setDropdown}
            onAiPreset={(text) => runCommand('ai-preset', text)}
            commentCount={commentCount}
            openCommentCount={openCommentCount}
            resolvedCommentCount={resolvedCommentCount}
            commentAtCaret={commentAtCaret}
            trackChanges={trackChanges}
            spellcheck={spellcheck}
            revisionDisplay={revisionDisplay}
            balloonMode={balloonMode}
            revisionCount={revisionCount}
            revisionFilter={revisionFilter}
            revisionAuthors={revisionAuthors}
            showReviewPane={showReviewPane}
            isProtected={isProtected}
            commentsAllowed={commentsAllowed}
            trackChangesForced={trackChangesForced}
            protectActive={protectActive}
          />
        ) : (
          <ViewTab
            hasDoc={hasDoc}
            zoom={zoom}
            darkPage={darkPage}
            showRuler={showRuler}
            showNav={showNav}
            viewMode={viewMode}
            readMode={readMode}
            showGrid={showGrid}
            splitView={splitView}
          />
        )}
      </div>
    </div>
  )
}

// memo + shallow-stable formatState/callback props: caret moves that change no
// displayed format skip re-rendering the whole ribbon
/** the ribbon's controls read their commands from here, so the tab components need no extra props */
const NO_SAVED_THEMES: readonly SavedTheme[] = []

function RibbonWithCommands(props: RibbonProps) {
  const {
    hasDoc,
    formatState,
    showAi,
    showMarks,
    showStylesPane,
    readOnlyRecommended,
    mergePreview,
    commentCount,
    mergeSource,
    windowsDetached,
    onCommand,
  } = props
  const hasComments = commentCount > 0
  const hasMergeSource = !!mergeSource
  const source = useMemo<RibbonCommandSource>(
    () => ({
      hasDoc,
      format: formatState,
      view: {
        showAi,
        showMarks,
        showStylesPane: !!showStylesPane,
        readOnlyRecommended: !!readOnlyRecommended,
        mergePreview: !!mergePreview,
        hasComments,
        hasMergeSource,
        windowsDetached: !!windowsDetached,
      },
      run: onCommand,
    }),
    [
      hasDoc,
      formatState,
      showAi,
      showMarks,
      showStylesPane,
      readOnlyRecommended,
      mergePreview,
      hasComments,
      hasMergeSource,
      windowsDetached,
      onCommand,
    ],
  )
  return (
    <RibbonCommandContext.Provider value={source}>
      <RibbonInner {...props} />
    </RibbonCommandContext.Provider>
  )
}

export const Ribbon = memo(RibbonWithCommands)
