/**
 * Shared pieces of the slides ribbon: the Props contract, common constants,
 * small layout components, and the RibbonTabCtx bundle handed to the
 * extracted tab components.
 */
import type { Dispatch, MouseEvent as ReactMouseEvent, ReactNode, SetStateAction } from 'react'
import type {
  AnimEffectKind,
  AnimTrigger,
  AnimationItem,
  EditChartOp,
  EditTableStyleOp,
  GetLayoutsResult,
  InsertKind,
  TransitionKind,
} from '../../shared/ipc'
import type { InkPenSettings, InkTool } from '../ink'
import type { WordArtPreset } from '@genoffice/ui'
import type { ChartPresetDef, IconDef, SmartArtDef } from '../insert-presets'
import type { SlideThemePreset } from '../themes'
import type { ChartStyleInfo } from '@genoffice/pptx-render'
import { useI18n, type StringKey } from '../i18n/locale'

export type InsertDropKey =
  'shapes' | 'icons' | 'chart' | 'smartart' | 'wordart' | 'zoom' | 'addanim'

export const BIG = 28

export type FormatCmd =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikeThrough'
  | 'superscript'
  | 'subscript'
  | 'removeFormat'
  | 'fontSizeUp'
  | 'fontSizeDown'

/** View modes: normal editing / outline / slide sorter / reading view. */
export type SlidesViewMode = 'normal' | 'outline' | 'sorter' | 'reading'

/** Font dropdown candidates (Western + common CJK/Traditional; the current font is inserted first when not in the list) */
export const FONT_FAMILIES = [
  'Calibri',
  'Calibri Light',
  'Arial',
  'Times New Roman',
  'Cambria',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Courier New',
  'Impact',
  '等线',
  '等线 Light',
  '宋体',
  '黑体',
  '微软雅黑',
  '楷体',
  '仿宋',
  'PingFang SC',
  'Noto Sans SC',
  'Noto Serif SC',
  'Yu Gothic',
  'Yu Mincho',
  'Meiryo',
  'MS Mincho',
  'Hiragino Sans',
  'Noto Sans JP',
  'Malgun Gothic',
  'Batang',
  'Apple SD Gothic Neo',
  'Microsoft JhengHei',
  'PMingLiU',
  'PingFang TC',
]

/** Font size dropdown candidates (pt, same ladder as PowerPoint) */
export const FONT_SIZES = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
]

/** Font color palette (applied with onMouseDown while editing, so the native picker doesn't steal focus and commit the edit) */
export const TEXT_COLORS = [
  '#000000',
  '#5A5A5A',
  '#FFFFFF',
  '#C43E1C',
  '#E97132',
  '#FFC000',
  '#4EA72E',
  '#0F9ED5',
  '#0A50A1',
  '#7030A0',
]

/** Thin dropdown chevron (replaces the ▾ text glyph) */
export function RbCaret() {
  return (
    <svg
      className="rb-caret"
      width="10"
      height="10"
      viewBox="0 0 24 24"
      fill="none"
      aria-hidden="true"
    >
      <path
        d="M5.5 9.25 12 15.75l6.5-6.5"
        stroke="currentColor"
        strokeWidth="2.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/** PowerPoint's canonical layout names → localized labels (the built-in set; unknown names show as-is) */
const LAYOUT_NAME_KEYS: Record<string, StringKey> = {
  'Title Slide': 'ribbonLayoutTitleSlide',
  'Title and Content': 'ribbonLayoutTitleAndContent',
  'Section Header': 'ribbonLayoutSectionHeader',
  'Two Content': 'ribbonLayoutTwoContent',
  'Title Only': 'ribbonLayoutTitleOnly',
  Blank: 'ribbonLayoutBlank',
}

/** Layout candidates with placeholder-sketch previews (new-slide dropdown + layout picker) */
export function LayoutList({
  layouts,
  size,
  onPick,
}: {
  layouts: GetLayoutsResult['layouts'] | null
  size: GetLayoutsResult['size'] | null
  onPick: (path: string) => void
}) {
  const { t } = useI18n()
  const W = 120
  const H = 68 // preview box (px)
  const cx = size?.cx || 9144000
  const cy = size?.cy || 5143500
  const list = layouts ?? []
  return (
    <div className="rb-layout-list">
      {list.map((lay) => {
        const key = LAYOUT_NAME_KEYS[lay.name]
        const name = key ? t(key) : lay.name
        return (
          <button
            key={lay.path}
            className="rb-layout-item"
            onClick={() => onPick(lay.path)}
            title={name}
          >
            <div className="rb-layout-preview">
              {lay.placeholders.map((ph, i) => (
                <div
                  key={i}
                  className="rb-layout-ph"
                  style={{
                    left: Math.round((ph.x / cx) * W),
                    top: Math.round((ph.y / cy) * H),
                    width: Math.max(8, Math.round((ph.cx / cx) * W)),
                    height: Math.max(6, Math.round((ph.cy / cy) * H)),
                  }}
                >
                  <span>
                    {ph.type === 'title' || ph.type === 'ctrTitle'
                      ? 'T'
                      : ph.type === 'body' || ph.type === 'obj'
                        ? '≡'
                        : ''}
                  </span>
                </div>
              ))}
            </div>
            <div className="rb-layout-name">{name}</div>
          </button>
        )
      })}
      {list.length === 0 && <div className="rb-layout-empty">{t('ribbonNoLayouts')}</div>}
    </div>
  )
}

/** Every ribbon popup that participates in "one popup at a time". */
export type RibbonPanelKey =
  | 'file'
  | 'color'
  | 'font'
  | 'size'
  | 'lineSpacing'
  | 'para'
  | 'layoutPick'
  | 'slideSize'
  | 'transparency'
  | 'pictureBorder'
  | 'table'
  | 'layout'
  | 'translate'
  | 'arrange'
  | 'insert'
  | 'chart'
  | 'collapse'
  | 'pen'
  | 'slideShow'

/** Ribbon popups are mutually exclusive: a trigger closes every sibling popup
 *  on mousedown, before its own click-toggle runs. A trigger rendered inside
 *  another popup (collapse flyout, paragraph panel) keeps its anchor open —
 *  closing it would unmount the popup being opened. */
export function closeSiblingPanels(
  e: ReactMouseEvent<HTMLElement>,
  closePanels: (keep: RibbonPanelKey[]) => void,
  own: RibbonPanelKey,
): void {
  const nested = e.currentTarget.closest('.rb-drop') != null
  closePanels(nested ? [own, 'collapse', 'para'] : [own])
}

export function Group({
  label,
  children,
  groupId,
  collapse,
}: {
  label: string
  children: ReactNode
  /** identity for width measurement + collapse bookkeeping */
  groupId?: string
  /** present on collapsible groups; `collapsed` switches to the dropdown form */
  collapse?: { collapsed: boolean; open: boolean; onToggle: () => void; icon: ReactNode }
}) {
  if (collapse?.collapsed) {
    return (
      <div className="ribbon-group" data-rbgroup={groupId}>
        <div className="ribbon-group-items">
          <div className="rb-drop-wrap">
            <button
              className={`rb-big ${collapse.open ? 'active' : ''}`}
              title={label}
              onMouseDown={(e) => e.stopPropagation()}
              onClick={collapse.onToggle}
            >
              <span className="rb-big-icon">
                {collapse.icon}
                <RbCaret />
              </span>
              <span>{label}</span>
            </button>
            {collapse.open && (
              <div className="rb-drop rb-collapse-panel" onMouseDown={(e) => e.stopPropagation()}>
                {children}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{label}</div>
      </div>
    )
  }
  return (
    <div className="ribbon-group" data-rbgroup={groupId}>
      <div className="ribbon-group-items">{children}</div>
      <div className="ribbon-group-label">{label}</div>
    </div>
  )
}

export interface Props {
  hasDoc: boolean
  /** True when no slide has real content — the one-click AI actions grey out then */
  deckEmpty: boolean
  /** Open file name (shown on the right of the tab row; the title bar row was removed) */
  dirty: boolean
  editing: boolean
  autoSave: boolean
  onAutoSaveChange: (on: boolean) => void
  onOpen: () => void
  onSave: () => void
  onUndo: () => void
  onRedo: () => void
  onSaveAs: () => void
  /** Export as PDF (hidden slides skipped) */
  onExportPdf: () => void
  onPrint: () => void
  /** Export as images (one PNG per page, hidden slides skipped) */
  onExportImages: () => void
  onFormat: (cmd: FormatCmd) => void
  zoom: number
  onZoom: (z: number | ((current: number) => number)) => void
  showThumbs: boolean
  onToggleThumbs: () => void
  aiOpen: boolean
  onToggleAi: () => void
  /** Push a preset instruction to the AI panel and expand it (autoRun executes immediately) */
  /** slideShot: attach the current slide's rendering so the model sees the page (AI Beautify) */
  onAiPreset: (text: string, opts?: { slideShot?: boolean }) => void
  /** Insert an element on the current page */
  onInsert: (kind: InsertKind) => void
  /** Shape gallery pick: enter canvas draw mode (crosshair; click = default size, drag = custom, Esc cancels) */
  onPickShape: (kind: InsertKind) => void
  /** Open the image picker dialog and insert into the current page */
  onInsertImage: () => void
  /** Set the page background solid color; allSlides=true applies to all pages */
  onBackground: (color: string, allSlides: boolean) => void
  /** Apply a built-in theme (colors + font scheme, applied to all pages) */
  onApplyTheme: (preset: SlideThemePreset) => void
  /** New blank slide (inherits the current page's layout background, empty content) */
  onAddSlide: () => void
  /** New slide with a given layout */
  onAddSlideWithLayout: (layoutPath: string) => void
  /** Add a section (before the current page, modeled on PowerPoint Home tab "Section") */
  onAddSection: () => void
  /** The current pptx's layout list (null = not loaded) */
  layouts: GetLayoutsResult['layouts'] | null
  /** Slide size (EMU) for normalizing layout previews */
  layoutSize: GetLayoutsResult['size'] | null
  formatOpen: boolean
  onToggleFormat: () => void
  hasSelection: boolean
  /** Whether the selection contains text-capable elements (text boxes/shapes/tables): font group availability outside editing (pictures/charts etc. grayed) */
  hasTextSelection: boolean
  canPaste: boolean
  onCopy: () => void
  onCut: () => void
  onPaste: () => void
  /** Format painter: whether a format has been copied */
  hasBrushFormat: boolean
  /** Format painter current mode (null=inactive; 'once'=single; 'continuous'=continuous) */
  brushMode: 'once' | 'continuous' | null
  /** Format painter button single click */
  onFormatBrushClick: () => void
  /** Format painter button double click (enters continuous mode) */
  onFormatBrushDoubleClick: () => void
  /** Editing: selection font color (execCommand) */
  onTextColor: (hex: string) => void
  /** Font group display: font/size of the current selection (editing) or aggregated from selected elements' text (null with no document) */
  curFontFamily: string | null
  curFontSizePt: number | null
  /** Mixed selection font sizes (shown as "min+") */
  curFontSizeMixed?: boolean
  /** Current bullet char of the selection for the bullet gallery ('' = no bullet; null = mixed/unknown, nothing highlighted) */
  curBulletChar: string | null
  /** Current paragraph alignment of the selection ('left' when unset; null = mixed/no text, nothing highlighted) */
  curAlign: 'left' | 'center' | 'right' | 'justify' | null
  /** Editing: change the selection's font / set size (pt) */
  onFontFamily: (family: string) => void
  onFontSize: (pt: number) => void
  /** Paragraph alignment: execCommand while editing, element-level op when elements are selected */
  onAlign: (align: 'left' | 'center' | 'right' | 'justify') => void
  /** Strikethrough: element-level toggle when selected but not editing (editing goes through onFormat) */
  onStrike: () => void
  /** B/I/U element-level toggle (when selected but not editing) */
  onTextToggle: (kind: 'bold' | 'italic' | 'underline') => void
  /** Element-level font color (selected but not editing; editing goes through onTextColor) */
  onElementTextColor: (hex: string) => void
  /** Open the find/replace panel (⌘F) */
  onFindReplace: () => void
  /** Animation "by paragraph" toggle (entrance effects split into one per paragraph) */
  animByParagraph: boolean
  onToggleAnimByParagraph: () => void
  /** Switch the current page's layout / reset layout (placeholders back in position) */
  onSetLayout: (layoutPath: string) => void
  onResetLayout: () => void
  /** Slide size (EMU) and the current size tag ('16:9'|'4:3'|null for display) */
  onSlideSize: (cx: number, cy: number) => void
  slideSizeKey: '16:9' | '4:3' | null
  /** Element-level paragraph format (bullets/numbering/line spacing) */
  onParagraphFormat: (patch: {
    bullet?: 'char' | 'number' | 'none'
    bulletChar?: string
    bulletHangEmu?: number
    bulletSizePct?: number
    bulletColor?: string
    lineSpacingPct?: number
    spaceBeforePt?: number
    spaceAfterPt?: number
    indentDelta?: 1 | -1
  }) => void
  onInsertTable: (rows: number, cols: number) => void
  /** Current page's transition effect (for display) */
  transition: TransitionKind
  /** Set the transition effect; allSlides=true applies to all pages */
  onTransition: (kind: TransitionKind, allSlides: boolean) => void
  // ── Animations tab ─────────────────────────────────────────────────────
  /** Selected shape's current animation effect (gallery highlight; null when no selection/no animation) */
  selectedAnimEffect: AnimEffectKind | null
  /** Animation bound to the timing controls (animation pane selection first, else the selected shape's last one) */
  timingAnim: AnimationItem | null
  /** Apply an animation to the selected shape (replacing its existing ones); 'none' removes all its animations */
  onApplyAnimation: (effect: AnimEffectKind | 'none') => void
  /** Hover preview: play the effect on the selection without applying it; null path = plain effect */
  onAnimHoverPreview: (effect: AnimEffectKind, motionPath?: string) => void
  onAnimHoverEnd: () => void
  /** Append one animation to the selected shape (stacking on existing ones, PowerPoint "Add Animation") */
  onAddAnimation: (effect: AnimEffectKind) => void
  /** Append one motion-path animation to the selected shape (moves along a preset path) */
  onApplyMotionPath: (path: string) => void
  /** Change timingAnim's trigger/duration/delay */
  onAnimTiming: (patch: { trigger?: AnimTrigger; durationMs?: number; delayMs?: number }) => void
  /** Animation pane (right side) toggle */
  animPaneOpen: boolean
  onToggleAnimPane: () => void
  /** Current page's animation count (preview button availability) */
  animCount: number
  /** Preview the current page's animations on the edit canvas */
  onAnimPreview: () => void
  /** Start the show (fromStart=true from the beginning, false from the current page) */
  onSlideShow: (fromStart: boolean) => void
  /** Start presenter view (single-window version: current page + next-page preview + notes + timer) */
  onPresenterView: (fromStart: boolean) => void
  /** Open the custom show management dialog (create/edit/play page subsets) */
  onCustomShow: () => void
  /** Start rehearsal timing (plays the show recording each page's dwell time; can be saved as auto-advance times afterwards) */
  onRehearse: () => void
  /** Whether the current page is hidden (hide-slide button display) */
  currentHidden: boolean
  /** Hide/unhide the current page */
  onToggleHidden: () => void
  /** Drawing tool (select = exit drawing) */
  inkTool: InkTool
  onInkTool: (tool: InkTool) => void
  inkPen: InkPenSettings
  onInkPen: (settings: InkPenSettings) => void
  inkHighlighter: InkPenSettings
  onInkHighlighter: (settings: InkPenSettings) => void
  /** Current page's stroke count (clear button availability) */
  inkCount: number
  /** Clear all ink on the current page */
  onInkClearAll: () => void
  /** View mode (normal/outline/slide sorter/reading) */
  viewMode: SlidesViewMode
  onViewMode: (mode: SlidesViewMode) => void
  /** Enter the slide master editing view */
  onSlideMaster: () => void
  /** Zoom to fit the window */
  onZoomFit: () => void
  showRuler: boolean
  onToggleRuler: () => void
  showGrid: boolean
  onToggleGrid: () => void
  showGuides: boolean
  onToggleGuides: () => void
  /** Notes pane (below the canvas, reads/writes pptx speaker notes) */
  showNotes: boolean
  onToggleNotes: () => void
  /** Comments pane (right side) */
  commentsOpen: boolean
  onToggleComments: () => void
  /** New comment: open the comments pane and focus the input */
  onNewComment: () => void
  /** Current page's comment count (button badge) */
  commentCount: number
  // ── Insert tab extensions ────────────────────────────────────────────────
  /** Insert an icon (rasterized to a PNG image) */
  onInsertIcon: (def: IconDef, color: string) => void
  /** Insert a chart (sample data, writes a chart part) */
  onInsertChart: (kind: ChartPresetDef['kind']) => void
  /** Insert SmartArt (simplified shape combination) */
  onInsertSmartArt: (def: SmartArtDef) => void
  /** Insert WordArt (preset-styled text box) */
  onInsertWordArt: (preset: WordArtPreset) => void
  /** Insert date-time / slide number text boxes (dynamic fields) */
  onInsertField: (type: 'datetime' | 'slidenum') => void
  /** Open the hyperlink dialog (requires a selected element) */
  onOpenLink: () => void
  /** Insert a Zoom link (button shape jumping to a given page) */
  onInsertZoom: (slideIndex: number) => void
  /** Document page count / current page (for the Zoom dropdown) */
  slideCount: number
  currentSlide: number
  /** Current slide's solid background color (undefined = gradient/image/none); syncs the Design tab swatch */
  currentBgColor?: string
  /** Open the header & footer dialog */
  onOpenHeaderFooter: () => void
  /** Open the equation dialog */
  onOpenEquation: () => void
  /** Open a dialog to insert video/audio */
  onInsertMedia: (kind: 'video' | 'audio') => void
  /** Open a dialog to insert a 3D model (embedded glb + poster placeholder) */
  onInsertModel3d: () => void
  /** Screen recording state (true = recording, button shows stop) */
  recording: boolean
  onToggleScreenRecord: () => void
  // ── Contextual tabs: table design / chart design / picture format ────────────────
  /** Current single-selection element type (undefined = none/multi-select; 'table'|'chart'|'picture' shows the contextual tab) */
  contextElementType?: 'table' | 'chart' | 'picture' | null
  /** Currently selected element sourceId (for contextual tab operation callbacks) */
  contextElementId?: string
  /** Current page index (for contextual tab operations) */
  contextSlideIndex?: number
  /** Chart created by this app (false = passthrough, chart design tab hidden) */
  /** Selected chart's current style (type highlight + chart element toggles display) */
  contextChartStyle?: ChartStyleInfo | null
  /** Theme-derived chart color schemes (falls back to the fixed palette when missing) */
  chartColorSchemes?: Array<{ key: string; label: string; colors: string[] }> | null
  /** Whether the selected picture supports background removal (audio/video poster frames etc. don't) */
  contextPictureCanCutout?: boolean
  /** Picture: enter crop mode */
  onPictureCrop?: () => void
  /** Crop mode is live — the Crop button shows its selected state */
  cropActive?: boolean
  /** Picture opacity (1 = opaque) */
  onPictureOpacity?: (opacity: number) => void
  /** Picture: enter cutout (background removal) mode */
  onPictureCutout?: () => void
  /** Selected picture's current border (null = none) */
  contextPictureStroke?: { color: string; widthPt: number; dashPreset?: string } | null
  /** Picture border (null clears it) */
  onPictureStroke?: (stroke: { color: string; widthPt: number; dash?: string } | null) => void
  /** Execute a table style operation */
  onEditTableStyle?: (op: Omit<EditTableStyleOp, 'slideIndex' | 'sourceId'>) => void
  /** Selected table's header-row/banded-rows current state (toggle display) */
  tableStyleFlags?: { firstRow: boolean; bandRow: boolean } | null
  /** Cell being edited in the selected table; shading applies to just this cell */
  tableActiveCell?: { row: number; col: number } | null
  /** Execute a chart edit operation */
  onEditChart?: (op: Omit<EditChartOp, 'slideIndex' | 'sourceId'>) => void
  /** Open the chart data edit dialog */
  onOpenChartDataDialog?: () => void
  /** Align/distribute selected elements (pure geometry, computed in the renderer then batchEditTransform) */
  onArrange?: (
    op:
      | 'left'
      | 'center-h'
      | 'right'
      | 'top'
      | 'center-v'
      | 'bottom'
      | 'distribute-h'
      | 'distribute-v',
  ) => void
  /** Mirror the selection horizontally/vertically (a:xfrm flipH/flipV) */
  onFlip?: (axis: 'h' | 'v') => void
  /** Whether distribute operations are allowed (≥3 selected elements) */
  canDistribute?: boolean
}

/** Ribbon locals + props handed to the extracted tab components; rebuilt every render. */
export interface RibbonTabCtx extends Pick<
  Props,
  | 'aiOpen'
  | 'brushMode'
  | 'canDistribute'
  | 'canPaste'
  | 'curBulletChar'
  | 'curAlign'
  | 'curFontFamily'
  | 'curFontSizeMixed'
  | 'curFontSizePt'
  | 'currentSlide'
  | 'deckEmpty'
  | 'editing'
  | 'formatOpen'
  | 'hasBrushFormat'
  | 'hasDoc'
  | 'hasSelection'
  | 'hasTextSelection'
  | 'layouts'
  | 'layoutSize'
  | 'onAddSection'
  | 'onAddSlide'
  | 'onAddSlideWithLayout'
  | 'onAiPreset'
  | 'onAlign'
  | 'onArrange'
  | 'onFlip'
  | 'onCopy'
  | 'onCut'
  | 'onElementTextColor'
  | 'onFindReplace'
  | 'onFontFamily'
  | 'onFontSize'
  | 'onFormat'
  | 'onFormatBrushClick'
  | 'onFormatBrushDoubleClick'
  | 'onInsert'
  | 'onPickShape'
  | 'onInsertChart'
  | 'onInsertField'
  | 'onInsertIcon'
  | 'onInsertImage'
  | 'onInsertMedia'
  | 'onInsertModel3d'
  | 'onInsertSmartArt'
  | 'onInsertTable'
  | 'onInsertWordArt'
  | 'onInsertZoom'
  | 'onNewComment'
  | 'onOpenEquation'
  | 'onOpenHeaderFooter'
  | 'onOpenLink'
  | 'onParagraphFormat'
  | 'onPaste'
  | 'onResetLayout'
  | 'onSetLayout'
  | 'onSlideShow'
  | 'onStrike'
  | 'onTextColor'
  | 'onTextToggle'
  | 'onToggleAi'
  | 'onToggleFormat'
  | 'onToggleScreenRecord'
  | 'recording'
  | 'slideCount'
  | 'zoom'
> {
  arrangeOpen: boolean
  closePanels: (keep: RibbonPanelKey[]) => void
  collapseOpen: string | null
  collapsedGroups: string[]
  colorOpen: boolean
  commitFontDraft: () => void
  commitSizeDraft: () => void
  fontDraft: string | null
  setFontDraft: Dispatch<SetStateAction<string | null>>
  dropBig: (
    key: InsertDropKey,
    icon: ReactNode,
    label: string,
    title: string,
    content: ReactNode,
    disabled?: boolean,
  ) => ReactNode
  fmtBtn: (cmd: FormatCmd, label: ReactNode, title: string, className?: string) => ReactNode
  fontOpen: boolean
  iconColor: string
  lastBulletColor: string
  lastColor: string
  layoutOpen: boolean
  layoutPickOpen: boolean
  lineSpacingOpen: boolean
  onCustomBulletColor: (value: string) => void
  onCustomTextColor: (value: string) => void
  paraOpen: boolean
  recentColors: string[]
  setArrangeOpen: Dispatch<SetStateAction<boolean>>
  setCollapseOpen: Dispatch<SetStateAction<string | null>>
  setColorOpen: Dispatch<SetStateAction<boolean>>
  setFontOpen: Dispatch<SetStateAction<boolean>>
  setIconColor: Dispatch<SetStateAction<string>>
  setInsertDrop: Dispatch<SetStateAction<InsertDropKey | null>>
  setLastColor: Dispatch<SetStateAction<string>>
  setLayoutOpen: Dispatch<SetStateAction<boolean>>
  setLayoutPickOpen: Dispatch<SetStateAction<boolean>>
  setLineSpacingOpen: Dispatch<SetStateAction<boolean>>
  setParaOpen: Dispatch<SetStateAction<boolean>>
  setSizeDraft: Dispatch<SetStateAction<string | null>>
  setSizeOpen: Dispatch<SetStateAction<boolean>>
  setSlideShowFromStart: Dispatch<SetStateAction<boolean>>
  setSlideShowOpen: Dispatch<SetStateAction<boolean>>
  setTableCustom: Dispatch<SetStateAction<{ r: number; c: number }>>
  setTableHover: Dispatch<SetStateAction<{ r: number; c: number }>>
  setTableOpen: Dispatch<SetStateAction<boolean>>
  sizeDraft: string | null
  sizeOpen: boolean
  /** Home-tab split button memory: the last chosen start mode (true = from beginning) */
  slideShowFromStart: boolean
  slideShowOpen: boolean
  t: ReturnType<typeof useI18n>['t']
  tableCustom: { r: number; c: number }
  tableHover: { r: number; c: number }
  tableOpen: boolean
}
