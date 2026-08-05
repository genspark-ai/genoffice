import { memo, useEffect, useRef, useState } from 'react'
import type { CSSProperties, ReactNode } from 'react'
import type { ChainedCommands, Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  selectedRect,
  setCellAttr,
  splitCell,
} from '@tiptap/pm/tables'
import type {
  Block,
  CustomNumberingLevel,
  DocDefaults,
  HeaderFooter,
  SectionSettings,
  SourceInfo,
  StyleInfo,
  ThemeColors,
  ThemeFonts,
} from '@genoffice/docx-engine'
import { HIGHLIGHT_CSS } from '../editor/extensions'
import { setParagraphDirection, setSelectionAlign } from '../editor/direction'
import { stepParagraphIndent } from '../editor/indent'
import { formatNumber } from '../editor/numbering'
import type { InkTool } from '../editor/ink'
import type { RibbonFormatState } from './ribbon-format-state'
import { setSelectedColumnWidth } from '../editor/table-sizing'
import { useI18n, type StringKey } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
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
} from './ribbon-tabs'
import { WRAP_OPTIONS } from './ContextMenu'
import { CropDialog, CutoutDialog } from './PictureDialogs'
import {
  GensparkMark,
  IconAlignCenter,
  IconAlignJustify,
  IconAlignLeft,
  IconAlignRight,
  IconBorderAll,
  IconBorderInner,
  IconBorderNone,
  IconBorderOuter,
  IconBullets,
  IconCaret,
  IconCellAlignBottom,
  IconCellAlignMiddle,
  IconCellAlignTop,
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
  IconPalette,
  IconPaste,
  IconPilcrow,
  IconRemoveBg,
  IconReplacePicture,
  IconRowDelete,
  IconRowInsertAbove,
  IconRowInsertBelow,
  IconShading,
  IconShrinkFont,
  IconSort,
  IconSplitCells,
  IconTableDelete,
} from './icons'
interface RibbonProps {
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
  /** document character styles, from ParsedDoc.styles (type === 'character') */
  styles?: Map<string, StyleInfo>
  /** document-wide text defaults from styles.xml */
  docDefaults?: DocDefaults
  /** Open the paragraph dialog (line-spacing rule / exact value entry lives there) */
  onParagraphDialog?: () => void
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
  onAddSource: (source: SourceInfo) => void
  /** TOC page-number backfill: docHeadings in document order → real page numbers (null when not computable) */
  headingPages?: () => number[] | null
  zoom: number
  onZoom: (zoom: number) => void
  /** compute zoom from the current window size (Word: page width / whole page) */
  onZoomFit: (mode: 'width' | 'page') => void
  darkCanvas: boolean
  onDarkCanvas: (v: boolean) => void
  onAiPreset: (instruction: string) => void
  /** external request (e.g. native menu Page Setup) to switch to a specific tab */
  tabRequest?: { tab: string; nonce: number } | null
  header: HeaderFooter | null
  onHeader: (next: HeaderFooter) => void
  onPageNumFormat: () => void
  onInsertField: (instr: string) => void
  footer: HeaderFooter | null
  onFooter: (next: HeaderFooter) => void
  /** Different first page (w:titlePg) */
  titlePg: boolean
  onTitlePg: (v: boolean) => void
  /** Different odd & even pages (settings.xml w:evenAndOddHeaders) */
  evenOddHf: boolean
  onEvenOddHf: (v: boolean) => void
  showMarks: boolean
  onShowMarks: (v: boolean) => void
  showRuler: boolean
  onShowRuler: (v: boolean) => void
  showNav: boolean
  onShowNav: (v: boolean) => void
  commentCount: number
  onShowComments: () => void
  /** Review → comments / revisions / compare / protection */
  canComment: boolean
  onNewComment: () => void
  trackChanges: boolean
  onTrackChanges: (on: boolean) => void
  revisionDisplay: RevisionDisplayMode
  onRevisionDisplay: (mode: RevisionDisplayMode) => void
  revisionCount: number
  onAcceptRevision: (all: boolean) => void
  onRejectRevision: (all: boolean) => void
  onGotoRevision: (dir: 1 | -1) => void
  isProtected: boolean
  onToggleProtection: () => void
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

interface PainterState {
  marks: Array<{ type: string; attrs: Record<string, unknown> }>
  para: Record<string, unknown>
}

function transformCase(s: string, mode: 'upper' | 'lower' | 'title' | 'sentence'): string {
  switch (mode) {
    case 'upper':
      return s.toUpperCase()
    case 'lower':
      return s.toLowerCase()
    case 'title':
      return s.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m) => m.toUpperCase())
    case 'sentence':
      return s.toLowerCase().replace(/(^\s*\p{L})|([.!?。!?]\s*\p{L})/gu, (m) => m.toUpperCase())
  }
}

// Word for Mac has no File ribbon tab: file actions live in the native menu
// bar (which we provide). Windows Word does have one, so keep it there.
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
/** shell tab mode: the tab strip above owns traffic lights / caption buttons */
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'

const TABS = (
  IS_MAC
    ? ['home', 'insert', 'draw', 'design', 'layout', 'references', 'review', 'view']
    : ['file', 'home', 'insert', 'draw', 'design', 'layout', 'references', 'review', 'view']
) as readonly string[]
const TABLE_TABS = ['tableDesign', 'tableLayout'] as const
const IMAGE_TABS = ['pictureFormat'] as const
type RibbonTab = (typeof TABS)[number] | (typeof TABLE_TABS)[number] | (typeof IMAGE_TABS)[number]

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
}

/** CSS px per cm at 96dpi (size inputs display in centimeters) */
const PX_PER_CM = 96 / 2.54

/** Word's picture size limits in cm (0.01"–22") */
export const PICTURE_CM_MIN = 0.03
export const PICTURE_CM_MAX = 55.87
export const clampPictureCm = (cm: number) => Math.min(PICTURE_CM_MAX, Math.max(PICTURE_CM_MIN, cm))

/** swallows every command when the document is read-only (protected / read mode) */
const NOOP_CHAIN = new Proxy(
  {},
  { get: (_t, prop) => (prop === 'run' ? () => false : () => NOOP_CHAIN) },
) as ChainedCommands

// Word's preset size list (also drives the grow/shrink font step buttons)
const FONT_SIZES = [
  5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
]

const THEME_COLORS: Array<{ nameKey: StringKey; hex: string }> = [
  { nameKey: 'ribbonColorWhite', hex: 'FFFFFF' },
  { nameKey: 'ribbonColorBlack', hex: '000000' },
  { nameKey: 'ribbonColorLightGray', hex: 'E7E6E6' },
  { nameKey: 'ribbonColorBlueGray', hex: '0E2841' },
  { nameKey: 'ribbonColorBlue', hex: '156082' },
  { nameKey: 'ribbonColorOrange', hex: 'E97132' },
  { nameKey: 'ribbonColorGreen', hex: '196B24' },
  { nameKey: 'ribbonColorSkyBlue', hex: '0F9ED5' },
  { nameKey: 'ribbonColorPurple', hex: 'A02B93' },
  { nameKey: 'ribbonColorLightGreenAlt', hex: '4EA72E' },
]

const THEME_COLOR_SHADES = [
  [
    'F2F2F2',
    '7F7F7F',
    'D0CECE',
    'DDEBF7',
    'DDEBF7',
    'FCE4D6',
    'E2F0D9',
    'DDEBF7',
    'E4DFEC',
    'E2F0D9',
  ],
  [
    'D9D9D9',
    '595959',
    'AEAAAA',
    'BDD7EE',
    '9DC3E6',
    'F8CBAD',
    'C6E0B4',
    '9DC3E6',
    'D9E1F2',
    'C6E0B4',
  ],
  [
    'BFBFBF',
    '3F3F3F',
    '757171',
    '8EA9DB',
    '5B9BD5',
    'F4B084',
    'A9D18E',
    '5B9BD5',
    'B4C6E7',
    'A9D18E',
  ],
  [
    'A6A6A6',
    '262626',
    '3A3838',
    '4472C4',
    '2E75B6',
    'C65911',
    '70AD47',
    '00B0F0',
    '8064A2',
    '70AD47',
  ],
  [
    '808080',
    '0D0D0D',
    '171616',
    '203864',
    '1F4E78',
    '843C0C',
    '375623',
    '0070C0',
    '5B315E',
    '385723',
  ],
]

/** Word standard colors */
const COLORS: Array<{ nameKey: StringKey; hex: string }> = [
  { nameKey: 'ribbonColorDarkRed', hex: 'C00000' },
  { nameKey: 'ribbonColorRed', hex: 'FF0000' },
  { nameKey: 'ribbonColorOrange', hex: 'FFC000' },
  { nameKey: 'ribbonColorYellow', hex: 'FFFF00' },
  { nameKey: 'ribbonColorLightGreen', hex: '92D050' },
  { nameKey: 'ribbonColorGreen', hex: '00B050' },
  { nameKey: 'ribbonColorLightBlue', hex: '00B0F0' },
  { nameKey: 'ribbonColorBlue', hex: '0070C0' },
  { nameKey: 'ribbonColorDarkBlue', hex: '002060' },
  { nameKey: 'ribbonColorPurple', hex: '7030A0' },
]

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

// ---- List library presets (bullets/numbering/multilevel): picking one creates a numbering definition ----

/** Bullet library: the chosen symbol is level 1; deeper levels rotate through ○/■ */
function bulletPresetLevels(glyph: string): CustomNumberingLevel[] {
  const rotation = [glyph, '○', '■']
  return Array.from({ length: 9 }, (_, i) => ({
    numFmt: 'bullet',
    lvlText: rotation[i % 3],
    indentLeft: 720 * (i + 1),
    hanging: 360,
  }))
}

/** Numbering library: the same format continues per level (%1 in the pattern becomes each level's counter) */
function numberPresetLevels(numFmt: string, pattern: string): CustomNumberingLevel[] {
  return Array.from({ length: 9 }, (_, i) => ({
    numFmt,
    lvlText: pattern.replace('%1', `%${i + 1}`),
    indentLeft: 720 * (i + 1),
    hanging: 360,
  }))
}

const BULLET_LIBRARY = ['•', '○', '■', '◆', '➢', '✦']

const NUMBER_LIBRARY: Array<{ numFmt: string; pattern: string }> = [
  { numFmt: 'decimal', pattern: '%1.' },
  { numFmt: 'decimal', pattern: '%1)' },
  { numFmt: 'upperRoman', pattern: '%1.' },
  { numFmt: 'upperLetter', pattern: '%1.' },
  { numFmt: 'lowerLetter', pattern: '%1)' },
  { numFmt: 'chineseCountingThousand', pattern: '%1、' },
]

const MULTILEVEL_LIBRARY: CustomNumberingLevel[][] = [
  // 1. / 1.1. / 1.1.1.
  Array.from({ length: 9 }, (_, i) => ({
    numFmt: 'decimal',
    lvlText: `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}.`,
    indentLeft: 720 * (i + 1),
    hanging: 432,
  })),
  // Chinese official-document hierarchy: numeral + comma / parenthesized numeral / 1.
  Array.from({ length: 9 }, (_, i): CustomNumberingLevel => {
    if (i === 0)
      return { numFmt: 'chineseCountingThousand', lvlText: '%1、', indentLeft: 720, hanging: 425 }
    if (i === 1)
      return { numFmt: 'chineseCountingThousand', lvlText: '(%2)', indentLeft: 1440, hanging: 425 }
    return { numFmt: 'decimal', lvlText: `%${i + 1}.`, indentLeft: 720 * (i + 1), hanging: 360 }
  }),
  // • / ○ / ■
  bulletPresetLevels('•'),
]

/** The level's number text when every level counter is 1 (gallery/dialog preview) */
function previewLevelText(levels: CustomNumberingLevel[], ilvl: number): string {
  const l = levels[ilvl]
  if (!l) return ''
  if (l.numFmt === 'bullet') return l.lvlText
  return l.lvlText.replace(/%(\d)/g, (_, n: string) =>
    formatNumber(1, levels[Number(n) - 1]?.numFmt ?? 'decimal'),
  )
}

const STYLE_GALLERY = [
  { key: 'p', labelKey: 'ribbonStyleNormal', className: 'style-normal' },
  { key: 'h1', labelKey: 'ribbonStyleHeading1', className: 'style-h1' },
  { key: 'h2', labelKey: 'ribbonStyleHeading2', className: 'style-h2' },
  { key: 'h3', labelKey: 'ribbonStyleHeading3', className: 'style-h3' },
] as const satisfies ReadonlyArray<{ key: string; labelKey: StringKey; className: string }>

/** Fallback character styles shown when the document has no character styles.
 * Emphasis = italic + accent color, Intense Emphasis = bold + accent color;
 * applied via the plain italic/bold marks plus docTextStyle color. */
const CHAR_STYLE_PRESETS: Array<{
  styleId: string
  labelKey: StringKey
  mark: 'italic' | 'bold'
  display: (accentHex: string) => CSSProperties
}> = [
  {
    styleId: '__preset_emphasis',
    labelKey: 'ribbonStyleEmphasis',
    mark: 'italic',
    display: (accentHex) => ({ fontStyle: 'italic', color: `#${accentHex}` }),
  },
  {
    styleId: '__preset_strong',
    labelKey: 'ribbonStyleIntenseEmphasis',
    mark: 'bold',
    display: (accentHex) => ({ fontWeight: 'bold', color: `#${accentHex}` }),
  },
]

/** Theme accent used by the preset character styles when the doc theme has none */
const DEFAULT_PRESET_ACCENT = '4472C4'

function findNumIdOfKind(blocks: Block[], kind: 'bullet' | 'ordered'): string | null {
  for (const b of blocks) {
    if (b.type === 'listItem' && b.list?.kind === kind) return b.list.numId
  }
  return null
}

function RibbonInner({
  quickActions,
  trailingActions,
  editor,
  formatState: fs,
  hasDoc,
  blocks,
  allocateNumId,
  createListDef,
  onParagraphDialog,
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
  pageColor,
  onPageColor,
  watermark,
  onWatermark,
  themeFonts,
  onThemeFonts,
  themeColors,
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
  onAddSource,
  headingPages,
  zoom,
  onZoom,
  onZoomFit,
  darkCanvas,
  onDarkCanvas,
  onAiPreset,
  tabRequest,
  header,
  onHeader,
  onPageNumFormat,
  onInsertField,
  footer,
  onFooter,
  titlePg,
  onTitlePg,
  evenOddHf,
  onEvenOddHf,
  showMarks,
  onShowMarks,
  showRuler,
  onShowRuler,
  showNav,
  onShowNav,
  commentCount,
  onShowComments,
  canComment,
  onNewComment,
  trackChanges,
  onTrackChanges,
  revisionDisplay,
  onRevisionDisplay,
  revisionCount,
  onAcceptRevision,
  onRejectRevision,
  onGotoRevision,
  isProtected,
  onToggleProtection,
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
  // The one-click AI actions need text to work on; grey them out on an empty document
  const docEmpty = !hasDoc || fs.docEmpty
  const [tab, setTab] = useState<RibbonTab>('home')
  const [dropdown, setDropdown] = useState<string | null>(null)
  const [penColor, setPenColor] = useState('C00000')
  const [penHighlight, setPenHighlight] = useState('yellow')
  const [painter, setPainter] = useState<PainterState | null>(null)
  const ribbonRef = useRef<HTMLDivElement>(null)
  const lastRegularTab = useRef<(typeof TABS)[number]>('home')
  const wasInTable = useRef(false)
  const wasInImage = useRef(false)
  /** Picture Format → remove background / crop dialogs */
  const [pictureDialog, setPictureDialog] = useState<'cutout' | 'crop' | null>(null)
  const [listDialog, setListDialog] = useState(false)

  useEffect(() => {
    if (tabRequest && (TABS as readonly string[]).includes(tabRequest.tab)) {
      const requested = tabRequest.tab as (typeof TABS)[number]
      lastRegularTab.current = requested
      setTab(requested)
      setDropdown(null)
    }
  }, [tabRequest])

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (ribbonRef.current && !ribbonRef.current.contains(e.target as Node)) setDropdown(null)
    }
    window.addEventListener('mousedown', close)
    return () => window.removeEventListener('mousedown', close)
  }, [])

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

  useEffect(() => {
    if (inTable && !wasInTable.current) {
      wasInTable.current = true
      setDropdown(null)
      setTab('tableLayout')
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

  /**
   * Replace the selected image's bytes (shared by Replace Picture / remove background / crop).
   * The original image's patch-save only supports size/alignment/wrap; swapping bytes must go
   * through the genImage new-image embed branch, so docxIndex is cleared (on save the old
   * block is treated as deleted, the new image is written at the same position, and
   * alignment/wrap are inherited from attributes).
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
      editor
        .chain()
        .focus()
        .updateAttributes('docProtected', {
          imageDataUrl: dataUrl,
          imageWidthPx: w,
          imageHeightPx: h,
          genImage: { base64: m[2], mime: m[1], widthPx: w, heightPx: h },
          docxIndex: null,
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

  /** Set the image display size proportionally (cm input; either side drives the other) */
  const setPictureSizeCm = (dim: 'w' | 'h', cm: number) => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const w = Number(attrs?.imageWidthPx)
    const h = Number(attrs?.imageHeightPx)
    if (attrs?.blockType !== 'image' || !w || !h || !(cm > 0)) return
    const px = clampPictureCm(cm) * PX_PER_CM
    const next =
      dim === 'w'
        ? {
            imageWidthPx: Math.max(1, Math.round(px)),
            imageHeightPx: Math.max(1, Math.round((px * h) / w)),
          }
        : {
            imageWidthPx: Math.max(1, Math.round((px * w) / h)),
            imageHeightPx: Math.max(1, Math.round(px)),
          }
    editor.chain().focus().updateAttributes('docProtected', next).run()
  }

  /** Reset to the image's natural size (shrunk to 620px when exceeding body width, matching insertion) */
  const resetPictureSize = async () => {
    if (!canEdit) return
    const attrs = editor.getAttributes('docProtected')
    const url = attrs?.imageDataUrl as string | null
    if (attrs?.blockType !== 'image' || !url) return
    try {
      const natural = await imageSizeOf(url)
      const scale = Math.min(1, 620 / natural.width)
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
    command(editor.state, editor.view.dispatch)
  }

  // ---- Table borders / vertical alignment / row height & column width ----
  const [borderColor, setBorderColor] = useState('000000')
  const [borderSz, setBorderSz] = useState(4) // 1/8 pt:4 = 0.5pt
  const sectionContentWidthPx = section
    ? Math.max(1, (section.pageWidth - section.marginLeft - section.marginRight) / 15)
    : 624
  const maxRowHeightCm = section
    ? (Math.max(1, section.pageHeight - section.marginTop - section.marginBottom) / 1440) * 2.54
    : 23.28

  type BorderSide = { style: string; szEighths?: number; color?: string }
  /** Apply borders to selected cells: all/outer/inner compute the four sides per cell from selection geometry; none clears explicitly */
  const applyCellBorders = (mode: 'all' | 'outer' | 'inner' | 'none') => {
    if (!canEdit || !isInTable(editor.state)) return
    editor.view.focus()
    const { state, view } = editor
    const rect = selectedRect(state)
    const solid: BorderSide = { style: 'single', szEighths: borderSz, color: borderColor }
    const none: BorderSide = { style: 'none' }
    let tr = state.tr
    const seen = new Set<number>()
    for (let row = rect.top; row < rect.bottom; row++) {
      for (let col = rect.left; col < rect.right; col++) {
        const cellPos = rect.map.map[row * rect.map.width + col]
        if (seen.has(cellPos)) continue
        seen.add(cellPos)
        const pos = rect.tableStart + cellPos
        const node = state.doc.nodeAt(pos)
        if (!node) continue
        const cellRect = rect.map.findCell(cellPos)
        const edge = {
          top: cellRect.top <= rect.top,
          bottom: cellRect.bottom >= rect.bottom,
          left: cellRect.left <= rect.left,
          right: cellRect.right >= rect.right,
        }
        const next: Record<string, BorderSide> = {
          ...((node.attrs.borders as Record<string, BorderSide> | null) ?? {}),
        }
        for (const side of ['top', 'bottom', 'left', 'right'] as const) {
          if (mode === 'all') next[side] = solid
          else if (mode === 'none') next[side] = none
          else if (mode === 'outer' && edge[side]) next[side] = solid
          else if (mode === 'inner' && !edge[side]) next[side] = solid
        }
        tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders: next })
      }
    }
    view.dispatch(tr)
  }

  /** Set row height for selected rows (cm; 0/empty = clear) */
  const applyRowHeight = (cm: number | null) => {
    if (!canEdit || !isInTable(editor.state)) return
    editor.view.focus()
    const { state, view } = editor
    const rect = selectedRect(state)
    const twips = cm && cm > 0 ? Math.round((Math.min(cm, maxRowHeightCm) / 2.54) * 1440) : null
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

  /** Set column width for selected columns (cm): writes the matching colwidth slot of every cell in the column */
  const applyColumnWidth = (cm: number | null) => {
    if (!canEdit || !isInTable(editor.state) || !cm || cm <= 0) return
    editor.view.focus()
    const px = Math.round((cm / 2.54) * 96)
    setSelectedColumnWidth(px, sectionContentWidthPx)(editor.state, editor.view.dispatch)
  }

  /** Current cell properties (echoed in the size inputs) */
  const activeCellInfo =
    fs.cellKey === null
      ? null
      : {
          key: fs.cellKey,
          heightCm: fs.cellHeightCm,
          widthCm: fs.cellWidthCm,
          vAlign: fs.cellVAlign,
        }

  const activeCharStyleId = fs.charStyleId
  const presetAccent = themeColors?.accent1?.trim().toUpperCase() || DEFAULT_PRESET_ACCENT

  /**
   * Character styles shown in the gallery.
   * Use doc's own character styles (type=character) if any, otherwise show
   * the two built-in presets so the gallery is never empty.
   */
  const charStyleItems: Array<{ key: string; label: string; previewStyle: CSSProperties }> =
    (() => {
      // Collect non-Hyperlink character styles from the document
      const docItems: Array<{ key: string; label: string; previewStyle: React.CSSProperties }> = []
      if (styles) {
        for (const [id, info] of styles) {
          if (info.type !== 'character') continue
          if (id === 'Hyperlink' || id === 'FollowedHyperlink' || id === 'DefaultParagraphFont')
            continue
          // Word rule: semiHidden and linked character shells ("Heading 1 Char") stay out of the style gallery
          if (info.semiHidden || info.linkedCharShell) continue
          const css: CSSProperties = {}
          if (info.display?.bold) css.fontWeight = 'bold'
          if (info.display?.italic) css.fontStyle = 'italic'
          if (info.display?.underline) css.textDecoration = 'underline'
          if (info.display?.color) css.color = `#${info.display.color}`
          docItems.push({ key: `char:${id}`, label: info.name, previewStyle: css })
        }
      }
      if (docItems.length > 0) return docItems
      // Fallback: built-in presets
      return CHAR_STYLE_PRESETS.map((p) => ({
        key: `char:${p.styleId}`,
        label: t(p.labelKey),
        previewStyle: p.display(presetAccent),
      }))
    })()

  // Presets carry no styleId (they apply plain italic/bold + accent color), so
  // detect their active state from the format state instead of charStyleId.
  const usingPresetFallback = charStyleItems[0]?.key === `char:${CHAR_STYLE_PRESETS[0].styleId}`
  const presetActive = (mark: 'italic' | 'bold'): boolean =>
    usingPresetFallback &&
    !activeCharStyleId &&
    (mark === 'italic' ? fs.italic : fs.bold) &&
    (fs.textColor ?? '').toUpperCase() === presetAccent
  const activeStyleKey =
    fs.headingLevel !== null
      ? `h${fs.headingLevel}`
      : activeCharStyleId
        ? `char:${activeCharStyleId}`
        : presetActive('bold')
          ? 'char:__preset_strong'
          : presetActive('italic')
            ? 'char:__preset_emphasis'
            : 'p'
  const currentSize = fs.fontSizePt
  const currentFont = fs.fontFamily
  // The "(Body)" entry means "no explicit run font — inherit the document's body
  // font", so it has to name that font rather than a fixed one: docDefaults is what
  // actually renders, the theme's minor font is what "+Body" resolves to.
  const bodyFontName = docDefaults?.asciiFont?.trim() || themeFonts?.minor?.trim() || 'Calibri'
  // computed unconditionally (not inside the dropdown render): cheap, and the
  // render-isolation test uses fontFamiliesFor calls as its render probe
  const fontFamilies = fontFamiliesFor(lang)
  // unset align follows the paragraph direction: start is left in LTR, right in RTL
  const activeAlign = fs.align ?? (fs.bidi ? 'right' : 'left')
  const activeSpacing = fs.lineSpacing

  /** merge new attrs into the docTextStyle mark, preserving the rest.
   * Only the patch is passed: setMark merges per existing mark and with the caret's
   * stored mark. Rebuilding from getAttributes read-back dropped the previous call's
   * value on a collapsed cursor (stored-mark changes don't re-render). */
  const setTextStyle = (patch: Record<string, unknown>) => {
    chain().setMark('docTextStyle', patch).run()
    setDropdown(null)
  }

  /** apply paragraph-level attrs to every block type in the selection */
  const setParaAttr = (attrs: Record<string, unknown>) => {
    if (sub) {
      // textbox paragraphs only support alignment; other keys are ignored
      chain().updateAttributes('docParagraph', attrs).run()
      setDropdown(null)
      return
    }
    let c = chain()
      .updateAttributes('docParagraph', attrs)
      .updateAttributes('docHeading', attrs)
      .updateAttributes('docListItem', attrs)
    // alignment also applies to selected images (w:jc on the image paragraph)
    if ('align' in attrs) {
      c = c.updateAttributes('docProtected', { imageAlign: attrs.align ?? null })
    }
    c.run()
    setDropdown(null)
  }

  const applyStyle = (key: string) => {
    if (key.startsWith('char:')) {
      const styleId = key.slice(5)
      const preset = CHAR_STYLE_PRESETS.find((p) => p.styleId === styleId)
      if (preset) {
        // Presets are plain italic/bold marks + accent color; toggle off when already active
        if (activeStyleKey === key) {
          chain().unsetMark(preset.mark).setMark('docTextStyle', { color: null }).run()
        } else {
          // switching presets must drop the other's mark, otherwise both stay
          // active and the gallery highlight sticks on the wrong card
          let c = chain()
          for (const p of CHAR_STYLE_PRESETS) if (p !== preset) c = c.unsetMark(p.mark)
          c.setMark(preset.mark).setMark('docTextStyle', { color: presetAccent }).run()
        }
        return
      }
      // Toggle: if already active, remove the mark; else set it
      if (activeCharStyleId === styleId) {
        chain().unsetMark('docTextStyle').run()
      } else {
        chain().setMark('docTextStyle', { styleId }).run()
      }
      return
    }
    if (sub) return // textboxes have no heading styles
    let c = chain()
    if (key === 'p') c = c.setNode('docParagraph')
    else c = c.setNode('docHeading', { level: Number(key.slice(1)) })
    // Word-like: applying a paragraph style sheds the runs' direct font/size/color.
    // Those render as inline span styles and would otherwise mask the style's look
    // entirely (the click would seem to do nothing on documents whose body runs
    // carry explicit rPr, common in CJK templates).
    c.command(({ tr }) => {
      const { from, to } = tr.selection
      let start = from
      let end = to
      tr.doc.nodesBetween(from, to, (node, pos) => {
        if (node.isTextblock) {
          start = Math.min(start, pos + 1)
          end = Math.max(end, pos + node.nodeSize - 1)
        }
      })
      const type = editor.schema.marks.docTextStyle
      const jobs: Array<{ from: number; to: number; attrs: Record<string, unknown> | null }> = []
      tr.doc.nodesBetween(start, end, (node, pos) => {
        if (!node.isText) return
        const m = node.marks.find((mm) => mm.type === type)
        if (!m) return
        if (m.attrs.color == null && m.attrs.sizeHalfPoints == null && m.attrs.font == null) return
        const attrs = { ...m.attrs, color: null, sizeHalfPoints: null, font: null }
        const keep = Object.values(attrs).some((v) => v !== null)
        jobs.push({
          from: Math.max(pos, start),
          to: Math.min(pos + node.nodeSize, end),
          attrs: keep ? attrs : null,
        })
      })
      for (const job of jobs) {
        tr.removeMark(job.from, job.to, type)
        if (job.attrs) tr.addMark(job.from, job.to, type.create(job.attrs))
      }
      return true
    }).run()
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

  /** Custom levels picked in the gallery/dialog → create a definition and apply it to the current paragraph */
  const applyListPreset = (levels: CustomNumberingLevel[]) => {
    if (sub) return
    const numId = createListDef?.(levels) ?? null
    if (!numId) return
    const kind = levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered'
    const ilvl = editor.isActive('docListItem')
      ? Number(editor.getAttributes('docListItem').ilvl) || 0
      : 0
    chain().setNode('docListItem', { kind, numId, ilvl }).run()
  }

  const changeIndent = (delta: 1 | -1) => {
    if (sub || !canEdit) return
    stepParagraphIndent(editor, delta)
  }

  const stepFontSize = (dir: 1 | -1) => {
    const idx = FONT_SIZES.findIndex((s) => s >= currentSize)
    let next: number
    if (dir === 1)
      next =
        FONT_SIZES[
          Math.min(
            idx === -1 ? FONT_SIZES.length : idx + (FONT_SIZES[idx] === currentSize ? 1 : 0),
            FONT_SIZES.length - 1,
          )
        ]
    else next = FONT_SIZES[Math.max(idx === -1 ? FONT_SIZES.length - 1 : idx - 1, 0)]
    setTextStyle({ sizeHalfPoints: Math.round(next * 2) })
  }

  const toggleVertAlign = (kind: 'superscript' | 'subscript') => {
    setTextStyle({ vertAlign: fs.vertAlign === kind ? null : kind })
  }

  /** format painter: pick up formatting now, apply to the next selection */
  const togglePainter = () => {
    if (painter) {
      setPainter(null)
      return
    }
    if (!canEdit) return
    const marks = editor.state.selection.$head
      .marks()
      .map((m) => ({ type: m.type.name, attrs: { ...m.attrs } }))
    setPainter({
      marks,
      para: {
        align: fs.align,
        lineSpacing: fs.lineSpacing,
        shadingFill: fs.shadingFill,
        borders: fs.paraBorders ?? null,
        bidi: fs.bidi,
      },
    })
  }

  useEffect(() => {
    if (!painter) return
    let selectingWithMouse = false
    let finished = false
    let keyboardTimer: ReturnType<typeof setTimeout> | null = null

    const applyFinalSelection = () => {
      if (finished || !editor.isEditable) return
      const { from, to } = editor.state.selection
      if (from === to) return
      finished = true
      if (keyboardTimer) clearTimeout(keyboardTimer)
      setPainter(null)
      let c = editor.chain().focus().unsetAllMarks()
      for (const m of painter.marks) c = c.setMark(m.type, m.attrs)
      c.updateAttributes('docParagraph', painter.para)
        .updateAttributes('docHeading', painter.para)
        .updateAttributes('docListItem', painter.para)
        .run()
    }

    const onMouseDown = (event: MouseEvent) => {
      if (!editor.view.dom.contains(event.target as globalThis.Node)) return
      selectingWithMouse = true
      if (keyboardTimer) clearTimeout(keyboardTimer)
    }
    const onMouseUp = () => {
      if (!selectingWithMouse) return
      selectingWithMouse = false
      requestAnimationFrame(applyFinalSelection)
    }
    const onSelectionUpdate = () => {
      if (selectingWithMouse || finished) return
      if (keyboardTimer) clearTimeout(keyboardTimer)
      keyboardTimer = setTimeout(applyFinalSelection, 180)
    }

    editor.view.dom.addEventListener('mousedown', onMouseDown)
    window.addEventListener('mouseup', onMouseUp)
    editor.on('selectionUpdate', onSelectionUpdate)
    return () => {
      if (keyboardTimer) clearTimeout(keyboardTimer)
      editor.view.dom.removeEventListener('mousedown', onMouseDown)
      window.removeEventListener('mouseup', onMouseUp)
      editor.off('selectionUpdate', onSelectionUpdate)
    }
  }, [painter, editor])

  const changeCase = (mode: 'upper' | 'lower' | 'title' | 'sentence') => {
    if (!canEdit) return
    const { from, to } = ed.state.selection
    if (from === to) return
    ed.chain()
      .focus()
      .command(({ state, tr }) => {
        state.doc.nodesBetween(from, to, (node, pos) => {
          if (!node.isText || !node.text) return
          const start = Math.max(from, pos)
          const end = Math.min(to, pos + node.nodeSize)
          const slice = node.text.slice(start - pos, end - pos)
          const next = transformCase(slice, mode)
          if (next !== slice) {
            tr.replaceWith(
              tr.mapping.map(start),
              tr.mapping.map(end),
              state.schema.text(next, node.marks),
            )
          }
        })
        return true
      })
      .run()
    setDropdown(null)
  }

  const clipboard = async (action: 'cut' | 'copy' | 'paste') => {
    if (action !== 'copy' && !canEdit) return
    if (action === 'paste') {
      const text = await navigator.clipboard.readText()
      if (text) chain().insertContent(text).run()
    } else {
      document.execCommand(action)
      ed.commands.focus()
    }
  }

  const markBtn = (name: string, active: boolean, title: string, label: ReactNode) => (
    <button
      className={`rb-icon ${active ? 'active' : ''}`}
      disabled={!canEdit}
      title={title}
      onClick={() => chain().toggleMark(name).run()}
    >
      {label}
    </button>
  )

  return (
    <div className="ribbon" ref={ribbonRef}>
      <div
        className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}`}
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
              <div className="file-menu">
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
            className={`ribbon-tab ${tab === tabName ? 'active' : ''}`}
            onClick={() => {
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
              className={`ribbon-tab ${tab === tableTab ? 'active' : ''}`}
              onClick={() => {
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
              className={`ribbon-tab ${tab === imageTab ? 'active' : ''}`}
              onClick={() => {
                setTab(imageTab)
                setDropdown(null)
              }}
            >
              {t(TAB_LABEL_KEYS[imageTab])}
            </button>
          ))}
        <span className="ribbon-tabs-spacer" />
        {trailingActions}
      </div>

      <div className="ribbon-body">
        {tab === 'pictureFormat' && inImage ? (
          <div className="table-ribbon-body">
            {/* ---- Adjust: remove background / crop / replace picture ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <button
                  className="rb-big"
                  disabled={!canEdit}
                  title={t('ribbonRemoveBgTip')}
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
                  title={t('ribbonCropTip')}
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
                  title={t('ribbonReplacePictureTip')}
                  onClick={() => void replacePicture()}
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
            {/* ---- Arrange: wrap text / align ---- */}
            <div className="table-tool-group">
              <div className="table-tool-row">
                <select
                  className="rb-select"
                  disabled={!canEdit}
                  title={t('ribbonWrapText')}
                  value={fs.imageWrap ?? ''}
                  onChange={(e) => {
                    if (!canEdit) return
                    editor
                      .chain()
                      .focus()
                      .updateAttributes('docProtected', { imageWrap: e.target.value || null })
                      .run()
                  }}
                >
                  {WRAP_OPTIONS.map((opt) => (
                    <option key={String(opt.value)} value={opt.value ?? ''}>
                      {t(opt.labelKey)}
                    </option>
                  ))}
                </select>
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
                    title={label}
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
            {/* ---- Size: height/width (cm, proportional) + reset ---- */}
            <div className="table-tool-group">
              <div
                className="table-tool-row table-size-inputs"
                key={`${fs.imageWidthPx ?? ''}x${fs.imageHeightPx ?? ''}`}
              >
                <label>
                  {t('ribbonPicHeight')}
                  <input
                    type="number"
                    min={PICTURE_CM_MIN}
                    max={PICTURE_CM_MAX}
                    step={0.1}
                    defaultValue={
                      fs.imageHeightPx !== null ? (fs.imageHeightPx / PX_PER_CM).toFixed(2) : ''
                    }
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const v = parseFloat((e.target as HTMLInputElement).value)
                        if (Number.isFinite(v) && v > 0) setPictureSizeCm('h', v)
                      }
                    }}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      const cur = fs.imageHeightPx !== null ? fs.imageHeightPx / PX_PER_CM : null
                      if (
                        Number.isFinite(v) &&
                        v > 0 &&
                        (cur === null || Math.abs(v - cur) > 0.01)
                      ) {
                        setPictureSizeCm('h', v)
                      }
                    }}
                  />
                  {t('ribbonCm')}
                </label>
                <label>
                  {t('ribbonPicWidth')}
                  <input
                    type="number"
                    min={PICTURE_CM_MIN}
                    max={PICTURE_CM_MAX}
                    step={0.1}
                    defaultValue={
                      fs.imageWidthPx !== null ? (fs.imageWidthPx / PX_PER_CM).toFixed(2) : ''
                    }
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const v = parseFloat((e.target as HTMLInputElement).value)
                        if (Number.isFinite(v) && v > 0) setPictureSizeCm('w', v)
                      }
                    }}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      const cur = fs.imageWidthPx !== null ? fs.imageWidthPx / PX_PER_CM : null
                      if (
                        Number.isFinite(v) &&
                        v > 0 &&
                        (cur === null || Math.abs(v - cur) > 0.01)
                      ) {
                        setPictureSizeCm('w', v)
                      }
                    }}
                  />
                  {t('ribbonCm')}
                </label>
              </div>
              <div className="table-tool-row">
                <button title={t('ribbonResetSizeTip')} onClick={() => void resetPictureSize()}>
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
                  title={t('ribbonRemoveTableStyleTip')}
                  onClick={() => chain().updateAttributes('docTable', { tblStyleId: null }).run()}
                >
                  <span className="table-style-card-grid plain" />
                  <span>{t('ribbonNoStyle')}</span>
                </button>
                {[...(styles?.values() ?? [])]
                  .filter((info) => info.type === 'table' && info.styleId !== 'TableNormal')
                  .slice(0, 8)
                  .map((info) => (
                    <button
                      key={info.styleId}
                      className="table-style-card"
                      title={t('ribbonApplyTableStyleTip', { name: info.name })}
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
              <div className="table-style-swatches">
                {['FFFFFF', 'D9EAF7', 'FFF2CC', 'E2F0D9', 'FCE4D6', 'E4DFEC'].map((hex) => (
                  <button
                    key={hex}
                    className="table-style-swatch"
                    title={t('ribbonCellShadingTip', { hex })}
                    style={{ background: `#${hex}` }}
                    onClick={() => runTableCommand(setCellAttr('fill', hex))}
                  />
                ))}
                <button
                  className="table-style-clear"
                  onClick={() => runTableCommand(setCellAttr('fill', null))}
                >
                  {t('ribbonNoShading')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupShading')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-grid table-tool-grid-four">
                <button title={t('ribbonAllBordersTip')} onClick={() => applyCellBorders('all')}>
                  <IconBorderAll />
                  {t('ribbonAllBorders')}
                </button>
                <button
                  title={t('ribbonOuterBordersTip')}
                  onClick={() => applyCellBorders('outer')}
                >
                  <IconBorderOuter />
                  {t('ribbonOuterBorders')}
                </button>
                <button
                  title={t('ribbonInnerBordersTip')}
                  onClick={() => applyCellBorders('inner')}
                >
                  <IconBorderInner />
                  {t('ribbonInnerBorders')}
                </button>
                <button title={t('ribbonClearBordersTip')} onClick={() => applyCellBorders('none')}>
                  <IconBorderNone />
                  {t('ribbonNoBorders')}
                </button>
              </div>
              <div className="table-tool-row table-border-opts">
                <input
                  type="color"
                  title={t('ribbonBorderColor')}
                  value={`#${borderColor}`}
                  onChange={(e) => setBorderColor(e.target.value.slice(1).toUpperCase())}
                />
                <select
                  title={t('ribbonBorderWidth')}
                  value={borderSz}
                  onChange={(e) => setBorderSz(Number(e.target.value))}
                >
                  <option value={4}>{t('ribbonPtValue', { n: 0.5 })}</option>
                  <option value={8}>{t('ribbonPtValue', { n: 1 })}</option>
                  <option value={12}>{t('ribbonPtValue', { n: 1.5 })}</option>
                  <option value={18}>{t('ribbonPtValue', { n: 2.25 })}</option>
                  <option value={24}>{t('ribbonPtValue', { n: 3 })}</option>
                </select>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupBorders')}</div>
            </div>
          </div>
        ) : tab === 'tableLayout' ? (
          <div className="table-ribbon-body">
            <div className="table-tool-group table-tool-delete">
              <button
                className="table-tool-button danger"
                onClick={() => runTableCommand(deleteTable)}
              >
                <IconTableDelete />
                {t('ribbonDeleteTable')}
              </button>
              <div className="ribbon-group-label">{t('ribbonGroupDelete')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
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
              <div className="ribbon-group-label">{t('ribbonGroupRowsCols')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                <button disabled={!fs.canMergeCells} onClick={() => runTableCommand(mergeCells)}>
                  <IconMergeCells />
                  {t('ribbonMergeCells')}
                </button>
                <button disabled={!fs.canSplitCell} onClick={() => runTableCommand(splitCell)}>
                  <IconSplitCells />
                  {t('ribbonSplitCells')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupMerge')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-grid table-tool-grid-two">
                <button onClick={() => runTableCommand(deleteRow)}>
                  <IconRowDelete />
                  {t('ribbonDeleteRow')}
                </button>
                <button onClick={() => runTableCommand(deleteColumn)}>
                  <IconColDelete />
                  {t('ribbonDeleteColumn')}
                </button>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupRowColOps')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div className="table-tool-row">
                {(
                  [
                    ['top', t('ribbonAlignTop'), IconCellAlignTop],
                    ['center', t('ribbonAlignMiddle'), IconCellAlignMiddle],
                    ['bottom', t('ribbonAlignBottom'), IconCellAlignBottom],
                  ] as const
                ).map(([v, label, Icon]) => (
                  <button
                    key={v}
                    className={
                      (activeCellInfo?.vAlign ?? 'top') === v
                        ? 'table-tool-button active'
                        : 'table-tool-button'
                    }
                    onClick={() => runTableCommand(setCellAttr('vAlign', v === 'top' ? null : v))}
                  >
                    <Icon />
                    {label}
                  </button>
                ))}
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupAlignment')}</div>
            </div>
            <div className="ribbon-sep" />
            <div className="table-tool-group">
              <div
                className="table-tool-row table-size-inputs"
                key={activeCellInfo?.key ?? 'nosel'}
              >
                <label>
                  {t('ribbonRowHeight')}
                  <input
                    type="number"
                    min={0}
                    max={maxRowHeightCm}
                    step={0.1}
                    placeholder={t('ribbonAuto')}
                    defaultValue={
                      activeCellInfo?.heightCm ? activeCellInfo.heightCm.toFixed(2) : ''
                    }
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const input = e.target as HTMLInputElement
                        const v = parseFloat(input.value)
                        const next =
                          Number.isFinite(v) && v > 0 ? Math.min(v, maxRowHeightCm) : null
                        if (next !== null) input.value = next.toFixed(2)
                        applyRowHeight(next)
                      }
                    }}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      const cur = activeCellInfo?.heightCm ?? null
                      const next = Number.isFinite(v) && v > 0 ? Math.min(v, maxRowHeightCm) : null
                      if (next !== null) e.target.value = next.toFixed(2)
                      if (next !== cur && (next !== null || cur !== null)) applyRowHeight(next)
                    }}
                  />
                  {t('ribbonCm')}
                </label>
                <label>
                  {t('ribbonColumnWidth')}
                  <input
                    type="number"
                    min={0}
                    max={(sectionContentWidthPx / 96) * 2.54}
                    step={0.1}
                    placeholder={t('ribbonAuto')}
                    defaultValue={activeCellInfo?.widthCm ? activeCellInfo.widthCm.toFixed(2) : ''}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const v = parseFloat((e.target as HTMLInputElement).value)
                        if (Number.isFinite(v) && v > 0) applyColumnWidth(v)
                      }
                    }}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      if (
                        Number.isFinite(v) &&
                        v > 0 &&
                        Math.abs(v - (activeCellInfo?.widthCm ?? 0)) > 0.01
                      ) {
                        applyColumnWidth(v)
                      }
                    }}
                  />
                  {t('ribbonCm')}
                </label>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupCellSize')}</div>
            </div>
          </div>
        ) : tab === 'home' ? (
          <>
            {/* ---- Genspark AI (first slot: entry + one-click AI actions) ---- */}
            <div className="ribbon-group">
              <div className="ribbon-group-items">
                <button
                  className={`rb-big ai-entry ${showAi ? 'active' : ''}`}
                  title={t('aiOpenAssistant')}
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
                  title={t('aiSummarizePrompt')}
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
                  title={t('aiPolishPrompt')}
                  onClick={() => onAiPreset(t('aiPolishPrompt'))}
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
                  title={t('aiTidyPrompt')}
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
                <button
                  className="rb-big"
                  disabled={!canEdit}
                  onClick={() => void clipboard('paste')}
                >
                  <span className="rb-big-icon">
                    <IconPaste size={28} />
                  </span>
                  <span>{t('ribbonPaste')}</span>
                </button>
                <div className="rb-col">
                  <button
                    className="rb-small"
                    disabled={!canEdit}
                    title={t('ribbonCutTip')}
                    onClick={() => void clipboard('cut')}
                  >
                    <IconCut />
                  </button>
                  <button
                    className="rb-small"
                    disabled={!hasDoc}
                    title={t('ribbonCopyTip')}
                    onClick={() => void clipboard('copy')}
                  >
                    <IconCopy />
                  </button>
                  <button
                    className={`rb-small ${painter ? 'active' : ''}`}
                    disabled={!canEdit || !!sub}
                    title={painter ? t('ribbonPainterActiveTip') : t('ribbonPainterTip')}
                    onClick={togglePainter}
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
                  <div className="rb-split-wrap">
                    <input
                      className="rb-select rb-font-family"
                      disabled={!canEdit}
                      key={`f:${currentFont}:${hasDoc}`}
                      defaultValue={currentFont}
                      placeholder={t('ribbonFontBodyNamed', { font: bodyFontName })}
                      title={t('ribbonFontFamilyTip')}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                      }}
                      onBlur={(e) => {
                        const v = e.target.value.trim()
                        if (v !== currentFont) setTextStyle({ font: v || null })
                      }}
                    />
                    <button
                      className="rb-caret rb-combo-caret"
                      disabled={!canEdit}
                      title={t('ribbonFontFamilyTip')}
                      onClick={() => setDropdown((v) => (v === 'fontFamily' ? null : 'fontFamily'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'fontFamily' && (
                      <div className="spacing-menu rb-font-family-menu">
                        <button
                          className={!currentFont ? 'active' : ''}
                          style={{ fontFamily: cssFontFamily(bodyFontName) }}
                          onClick={() => setTextStyle({ font: null })}
                        >
                          {t('ribbonFontBodyNamed', { font: bodyFontName })}
                        </button>
                        {fontFamilies
                          .filter((f) => f !== bodyFontName)
                          .map((f) => (
                            <button
                              key={f}
                              className={f === currentFont ? 'active' : ''}
                              style={{ fontFamily: cssFontFamily(f) }}
                              onClick={() => setTextStyle({ font: f })}
                            >
                              {f}
                            </button>
                          ))}
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <input
                      className="rb-select rb-font-size"
                      type="number"
                      min={1}
                      max={1638}
                      step={0.5}
                      disabled={!canEdit}
                      key={`s:${currentSize}:${hasDoc}`}
                      defaultValue={currentSize}
                      title={t('ribbonFontSizeTip')}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                      }}
                      onBlur={(e) => {
                        const v = Number(e.target.value)
                        if (!Number.isFinite(v) || v <= 0) return
                        const half = Math.round(Math.min(1638, Math.max(1, v)) * 2)
                        if (half !== Math.round(currentSize * 2))
                          setTextStyle({ sizeHalfPoints: half })
                      }}
                    />
                    <button
                      className="rb-caret rb-combo-caret"
                      disabled={!canEdit}
                      title={t('ribbonFontSizeTip')}
                      onClick={() => setDropdown((v) => (v === 'fontSize' ? null : 'fontSize'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'fontSize' && (
                      <div className="spacing-menu rb-font-size-menu">
                        {FONT_SIZES.map((s) => (
                          <button
                            key={s}
                            className={
                              Math.round(s * 2) === Math.round(currentSize * 2) ? 'active' : ''
                            }
                            onClick={() => setTextStyle({ sizeHalfPoints: Math.round(s * 2) })}
                          >
                            {s}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    title={t('ribbonGrowFont')}
                    onClick={() => stepFontSize(1)}
                  >
                    <IconGrowFont />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    title={t('ribbonShrinkFont')}
                    onClick={() => stepFontSize(-1)}
                  >
                    <IconShrinkFont />
                  </button>
                  <span className="rb-mini-sep" />
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon"
                      disabled={!canEdit}
                      title={t('ribbonChangeCase')}
                      onClick={() => setDropdown((v) => (v === 'case' ? null : 'case'))}
                    >
                      Aa
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'case' && (
                      <div className="spacing-menu case-menu">
                        <button onClick={() => changeCase('sentence')}>
                          {t('ribbonCaseSentence')}
                        </button>
                        <button onClick={() => changeCase('lower')}>{t('ribbonCaseLower')}</button>
                        <button onClick={() => changeCase('upper')}>{t('ribbonCaseUpper')}</button>
                        <button onClick={() => changeCase('title')}>{t('ribbonCaseTitle')}</button>
                      </div>
                    )}
                  </div>
                  <button
                    className="rb-icon"
                    disabled={!canEdit}
                    title={t('ribbonClearFormatting')}
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
                    className={`rb-icon rb-script ${fs.vertAlign === 'subscript' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonSubscript')}
                    onClick={() => toggleVertAlign('subscript')}
                  >
                    x<sub>2</sub>
                  </button>
                  <button
                    className={`rb-icon rb-script ${fs.vertAlign === 'superscript' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonSuperscript')}
                    onClick={() => toggleVertAlign('superscript')}
                  >
                    x<sup>2</sup>
                  </button>
                  <span className="rb-mini-sep" />
                  {/* highlight: main button applies pen color, caret opens palette */}
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon rb-color-btn ${fs.highlight ? 'active' : ''}`}
                      disabled={!canEdit}
                      title={t('ribbonTextHighlightColor')}
                      onClick={() =>
                        setTextStyle({
                          highlight: fs.highlight === penHighlight ? null : penHighlight,
                        })
                      }
                    >
                      <IconHighlight />
                      <span
                        className="rb-color-bar"
                        style={{ background: HIGHLIGHT_CSS[penHighlight] }}
                      />
                    </button>
                    <button
                      className="rb-caret rb-color-caret"
                      disabled={!canEdit}
                      onClick={() => setDropdown((v) => (v === 'highlight' ? null : 'highlight'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'highlight' && (
                      <div className="color-palette color-palette-highlight color-palette-highlight-word">
                        <div className="color-section-title color-highlight-title">
                          {t('ribbonHighlightColors')}
                        </div>
                        <div className="color-highlight-grid">
                          {HIGHLIGHTS.map((h) => (
                            <button
                              key={h}
                              className={`color-swatch color-highlight-swatch ${fs.highlight === h ? 'selected' : ''}`}
                              title={h}
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
                      title={t('ribbonFontColor')}
                      onClick={() =>
                        setTextStyle({ color: penColor === '000000' ? null : penColor })
                      }
                    >
                      <span className="rb-color-a">A</span>
                      <span className="rb-color-bar" style={{ background: `#${penColor}` }} />
                    </button>
                    <button
                      className="rb-caret rb-color-caret"
                      disabled={!canEdit}
                      onClick={() => setDropdown((v) => (v === 'color' ? null : 'color'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'color' && (
                      <div className="color-palette color-palette-word">
                        <button
                          className={`color-automatic ${!fs.textColor ? 'selected' : ''}`}
                          onClick={() => {
                            setPenColor('000000')
                            setTextStyle({ color: null })
                          }}
                        >
                          {t('ribbonAutomatic')}
                        </button>
                        <div className="color-section-title">{t('ribbonThemeColorsSection')}</div>
                        <div className="color-theme-base">
                          {THEME_COLORS.map((c) => (
                            <button
                              key={c.hex}
                              className={`color-swatch color-swatch-large ${fs.textColor === c.hex ? 'selected' : ''}`}
                              title={t(c.nameKey)}
                              style={{ background: `#${c.hex}` }}
                              onClick={() => {
                                setPenColor(c.hex)
                                setTextStyle({ color: c.hex === '000000' ? null : c.hex })
                              }}
                            />
                          ))}
                        </div>
                        <div className="color-theme-shades">
                          {THEME_COLOR_SHADES.flatMap((row, rowIndex) =>
                            row.map((hex, columnIndex) => (
                              <button
                                key={`${rowIndex}-${columnIndex}-${hex}`}
                                className={`color-swatch color-swatch-large ${fs.textColor === hex ? 'selected' : ''}`}
                                title={t('ribbonThemeColorShadeTip', {
                                  r: rowIndex + 1,
                                  c: columnIndex + 1,
                                })}
                                style={{ background: `#${hex}` }}
                                onClick={() => {
                                  setPenColor(hex)
                                  setTextStyle({ color: hex })
                                }}
                              />
                            )),
                          )}
                        </div>
                        <div className="color-section-title color-standard-title">
                          {t('ribbonStandardColors')}
                        </div>
                        <div className="color-standard-row">
                          {COLORS.map((c) => (
                            <button
                              key={c.hex}
                              className={`color-swatch color-swatch-large ${fs.textColor === c.hex ? 'selected' : ''}`}
                              title={t(c.nameKey)}
                              style={{ background: `#${c.hex}` }}
                              onClick={() => {
                                setPenColor(c.hex)
                                setTextStyle({ color: c.hex })
                              }}
                            />
                          ))}
                        </div>
                        <label className="color-more">
                          <span className="color-more-icon">
                            <IconPalette size={16} />
                          </span>
                          {t('ribbonMoreColors')}
                          <input
                            type="color"
                            value={`#${penColor}`}
                            onChange={(event) => {
                              const hex = event.currentTarget.value.slice(1).toUpperCase()
                              setPenColor(hex)
                              setTextStyle({ color: hex })
                            }}
                          />
                        </label>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupFont')}</div>
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
                      title={t('ribbonBullets')}
                      onClick={() => toggleList('bullet')}
                    >
                      <IconBullets />
                    </button>
                    <button
                      className="rb-caret"
                      disabled={!canEdit || !!sub}
                      title={t('ribbonBullets')}
                      onClick={() => setDropdown((v) => (v === 'bulletLib' ? null : 'bulletLib'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'bulletLib' && (
                      <div className="layout-menu list-gallery">
                        {BULLET_LIBRARY.map((glyph) => (
                          <button
                            key={glyph}
                            className="list-gallery-card"
                            onClick={() => {
                              applyListPreset(bulletPresetLevels(glyph))
                              setDropdown(null)
                            }}
                          >
                            {glyph}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.listOrdered ? 'active' : ''}`}
                      disabled={!canEdit || !!sub}
                      title={t('ribbonNumbering')}
                      onClick={() => toggleList('ordered')}
                    >
                      <IconNumbered />
                    </button>
                    <button
                      className="rb-caret"
                      disabled={!canEdit || !!sub}
                      title={t('ribbonNumbering')}
                      onClick={() => setDropdown((v) => (v === 'numberLib' ? null : 'numberLib'))}
                    >
                      <IconCaret />
                    </button>
                    {dropdown === 'numberLib' && (
                      <div className="layout-menu list-gallery">
                        {NUMBER_LIBRARY.map((n, i) => {
                          const levels = numberPresetLevels(n.numFmt, n.pattern)
                          return (
                            <button
                              key={i}
                              className="list-gallery-card"
                              onClick={() => {
                                applyListPreset(levels)
                                setDropdown(null)
                              }}
                            >
                              {[1, 2, 3]
                                .map((v) => n.pattern.replace('%1', formatNumber(v, n.numFmt)))
                                .join(' ')}
                            </button>
                          )
                        })}
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className="rb-icon"
                      disabled={!canEdit || !!sub}
                      title={t('ribbonMultilevelTip')}
                      onClick={() => setDropdown((v) => (v === 'multiLib' ? null : 'multiLib'))}
                    >
                      <IconMultilevel />
                    </button>
                    {dropdown === 'multiLib' && (
                      <div className="layout-menu list-gallery list-gallery-multi">
                        {MULTILEVEL_LIBRARY.map((levels, i) => (
                          <button
                            key={i}
                            className="list-gallery-card list-gallery-card-multi"
                            onClick={() => {
                              applyListPreset(levels)
                              setDropdown(null)
                            }}
                          >
                            {[0, 1, 2].map((lvl) => (
                              <span key={lvl} style={{ paddingLeft: lvl * 10 }}>
                                {previewLevelText(levels, lvl)} ———
                              </span>
                            ))}
                          </button>
                        ))}
                        <button
                          className="list-gallery-define"
                          onClick={() => {
                            setListDialog(true)
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
                    title={t('ribbonDecreaseIndent')}
                    onClick={() => changeIndent(-1)}
                  >
                    <IconIndentDec />
                  </button>
                  <button
                    className="rb-icon"
                    disabled={!canEdit || !!sub}
                    title={t('ribbonIncreaseIndent')}
                    onClick={() => changeIndent(1)}
                  >
                    <IconIndentInc />
                  </button>
                  <span className="rb-mini-sep" />
                  <button
                    className="rb-icon"
                    disabled
                    title={t('ribbonNotSupportedSuffix', { label: t('ribbonSort') })}
                  >
                    <IconSort />
                  </button>
                  <button
                    className={`rb-icon ${showMarks ? 'active' : ''}`}
                    disabled={!hasDoc}
                    title={t('ribbonShowMarks')}
                    onClick={() => onShowMarks(!showMarks)}
                  >
                    <IconPilcrow />
                  </button>
                </div>
                <div className="rb-row">
                  <button
                    className={`rb-icon ${activeAlign === 'left' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonAlignLeftTip')}
                    onClick={() => setSelectionAlign(ed, 'left')}
                  >
                    <IconAlignLeft />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'center' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonAlignCenterTip')}
                    onClick={() => setSelectionAlign(ed, 'center')}
                  >
                    <IconAlignCenter />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'right' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonAlignRightTip')}
                    onClick={() => setSelectionAlign(ed, 'right')}
                  >
                    <IconAlignRight />
                  </button>
                  <button
                    className={`rb-icon ${activeAlign === 'justify' ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={t('ribbonJustifyTip')}
                    onClick={() => setSelectionAlign(ed, 'justify')}
                  >
                    <IconAlignJustify />
                  </button>
                  <span className="rb-mini-sep" />
                  <button
                    className={`rb-icon ${!fs.bidi ? 'active' : ''}`}
                    disabled={!canEdit || !!sub}
                    title={t('ribbonDirLtrTip')}
                    onClick={() => setParagraphDirection(editor, 'ltr')}
                  >
                    <IconDirLtr />
                  </button>
                  <button
                    className={`rb-icon ${fs.bidi ? 'active' : ''}`}
                    disabled={!canEdit || !!sub}
                    title={t('ribbonDirRtlTip')}
                    onClick={() => setParagraphDirection(editor, 'rtl')}
                  >
                    <IconDirRtl />
                  </button>
                  <span className="rb-mini-sep" />
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${activeSpacing ? 'active' : ''}`}
                      disabled={!canEdit}
                      title={t('ribbonLineSpacing')}
                      onClick={() => setDropdown((v) => (v === 'spacing' ? null : 'spacing'))}
                    >
                      <IconLineSpacing />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'spacing' && (
                      <div className="spacing-menu">
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
                        <button
                          onClick={() =>
                            setParaAttr({ lineSpacing: null, lineRule: null, lineRawTwips: null })
                          }
                        >
                          {t('ribbonDefault')}
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
                      title={t('ribbonParagraphShading')}
                      onClick={() => setDropdown((v) => (v === 'shading' ? null : 'shading'))}
                    >
                      <IconShading />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'shading' && (
                      <div className="color-palette">
                        {COLORS.map((c) => (
                          <button
                            key={c.hex}
                            className="color-swatch"
                            style={{ background: `#${c.hex}` }}
                            title={t(c.nameKey)}
                            onClick={() => setParaAttr({ shadingFill: c.hex })}
                          />
                        ))}
                        <button
                          className="color-clear"
                          onClick={() => setParaAttr({ shadingFill: null })}
                        >
                          {t('ribbonNoShading')}
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="rb-split-wrap">
                    <button
                      className={`rb-icon ${fs.paraBorders ? 'active' : ''}`}
                      disabled={!canEdit}
                      title={t('ribbonParagraphBorders')}
                      onClick={() => setDropdown((v) => (v === 'borders' ? null : 'borders'))}
                    >
                      <IconBorderAll />
                      <span className="rb-caret-inline">
                        <IconCaret />
                      </span>
                    </button>
                    {dropdown === 'borders' && (
                      <div className="spacing-menu borders-menu">
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
            <div className="ribbon-group">
              <div className="ribbon-group-items style-gallery">
                {STYLE_GALLERY.map((s) => (
                  <button
                    key={s.key}
                    className={`style-card ${activeStyleKey === s.key ? 'active' : ''}`}
                    disabled={!canEdit || !!sub}
                    onClick={() => applyStyle(s.key)}
                  >
                    <span className={`style-card-preview ${s.className}`}>
                      {t('ribbonStylePreview')}
                    </span>
                    <span className="style-card-label">{t(s.labelKey)}</span>
                  </button>
                ))}
                {/* all visible character styles render inline: the styles panel
                    that used to catch the overflow is gone */}
                {charStyleItems.map((s) => (
                  <button
                    key={s.key}
                    className={`style-card style-card-char ${activeStyleKey === s.key ? 'active' : ''}`}
                    disabled={!canEdit}
                    title={s.label}
                    onClick={() => applyStyle(s.key)}
                  >
                    <span className="style-card-preview" style={s.previewStyle}>
                      Aa
                    </span>
                    <span className="style-card-label">{s.label}</span>
                  </button>
                ))}
              </div>
              <div className="ribbon-group-label">{t('ribbonGroupStyles')}</div>
            </div>
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
            titlePg={titlePg}
            onTitlePg={onTitlePg}
            evenOddHf={evenOddHf}
            onEvenOddHf={onEvenOddHf}
            commentCount={commentCount}
            onShowComments={onShowComments}
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
            onShowComments={onShowComments}
            canComment={canComment}
            onNewComment={onNewComment}
            trackChanges={trackChanges}
            onTrackChanges={onTrackChanges}
            revisionDisplay={revisionDisplay}
            onRevisionDisplay={onRevisionDisplay}
            revisionCount={revisionCount}
            onAcceptRevision={onAcceptRevision}
            onRejectRevision={onRejectRevision}
            onGotoRevision={onGotoRevision}
            isProtected={isProtected}
            onToggleProtection={onToggleProtection}
            onCompare={onCompare}
          />
        ) : (
          <ViewTab
            hasDoc={hasDoc}
            filePath={filePath}
            zoom={zoom}
            onZoom={onZoom}
            onZoomFit={onZoomFit}
            showAi={showAi}
            onToggleAi={onToggleAi}
            darkCanvas={darkCanvas}
            onDarkCanvas={onDarkCanvas}
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
      {listDialog && (
        <ListDefineDialog
          onApply={(levels) => {
            setListDialog(false)
            applyListPreset(levels)
          }}
          onClose={() => setListDialog(false)}
        />
      )}
    </div>
  )
}

// memo + shallow-stable formatState/callback props: caret moves that change no
// displayed format skip re-rendering the whole ribbon
export const Ribbon = memo(RibbonInner)

// ---- Define New Multilevel List dialog ----

const LIST_NUM_FMTS = [
  'decimal',
  'bullet',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'chineseCountingThousand',
] as const

const LIST_FMT_SAMPLES: Record<string, string> = {
  decimal: '1, 2, 3',
  bullet: '● ○ ■',
  lowerLetter: 'a, b, c',
  upperLetter: 'A, B, C',
  lowerRoman: 'i, ii, iii',
  upperRoman: 'I, II, III',
  chineseCountingThousand: '一, 二, 三',
}

const TWIPS_PER_CM = 567

function ListDefineDialog({
  onApply,
  onClose,
}: {
  onApply: (levels: CustomNumberingLevel[]) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [levels, setLevels] = useState<CustomNumberingLevel[]>(MULTILEVEL_LIBRARY[0])
  const update = (i: number, patch: Partial<CustomNumberingLevel>) =>
    setLevels((ls) => ls.map((l, k) => (k === i ? { ...l, ...patch } : l)))
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal list-define-modal">
        <h2>{t('ribbonDefineNewList')}</h2>
        <div className="list-define-head">
          <span>{t('ribbonListLevel')}</span>
          <span>{t('ribbonListNumStyle')}</span>
          <span>{t('ribbonListFormatText')}</span>
          <span>{t('ribbonListIndentCm')}</span>
          <span>{t('ribbonListPreviewCol')}</span>
        </div>
        <div className="list-define-rows">
          {levels.map((l, i) => (
            <div key={i} className="list-define-row">
              <span>{i + 1}</span>
              <select
                value={l.numFmt}
                onChange={(e) => {
                  const numFmt = e.target.value
                  update(i, {
                    numFmt,
                    lvlText:
                      numFmt === 'bullet'
                        ? '•'
                        : l.lvlText.includes('%')
                          ? l.lvlText
                          : `%${i + 1}.`,
                  })
                }}
              >
                {LIST_NUM_FMTS.map((f) => (
                  <option key={f} value={f}>
                    {LIST_FMT_SAMPLES[f]}
                  </option>
                ))}
              </select>
              <input value={l.lvlText} onChange={(e) => update(i, { lvlText: e.target.value })} />
              <input
                type="number"
                step="0.25"
                min="0"
                value={+(l.indentLeft / TWIPS_PER_CM).toFixed(2)}
                onChange={(e) =>
                  update(i, {
                    indentLeft: Math.max(
                      0,
                      Math.round(parseFloat(e.target.value || '0') * TWIPS_PER_CM),
                    ),
                  })
                }
              />
              <span className="list-define-preview" style={{ paddingLeft: Math.min(i * 8, 48) }}>
                {previewLevelText(levels, i)}
              </span>
            </div>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button onClick={() => onApply(levels)}>{t('ribbonOk')}</button>
        </div>
      </div>
    </div>
  )
}
