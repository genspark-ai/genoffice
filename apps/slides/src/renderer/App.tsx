import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type {
  GroupRenderNode,
  RenderFill,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
  ChartRenderNode,
  PictureRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'
import type {
  AiSettings,
  AnimEffectKind,
  AnimTrigger,
  AnimationItem,
  AttachmentMeta,
  EditChartOp,
  EditParagraph,
  EditTableStyleOp,
  GetLayoutsResult,
  GradientFillSpec,
  InsertKind,
  LinkTargetOp,
  MasterPartItem,
  PasteSlideMode,
  SectionInfo,
  SlideComment,
  TransitionKind,
} from '../shared/ipc'
import { SlideCanvas, selectionChromeColor } from './SlideCanvas'
import { ZOOM_PREVIEW_EVENT } from './zoom-preview'
import { createWheelPager } from './wheel-page-flip'
import type { DrawRect } from './draw-shape'
import { SlideThumb } from './SlideThumb'
import { MasterView } from './MasterView'
import { TextEditOverlay, firstFontFamily, liveAlign, liveBulletChar } from './TextEditOverlay'
import { CropOverlay } from './CropOverlay'
import { createImageLoader } from './image-loader'
import { syncPrivateFonts } from './doc-fonts'
import { InkOverlay } from './InkOverlay'
import { inkNodesOf, type InkPenSettings, type InkStroke, type InkTool } from './ink'
import type { SlideThemePreset } from './themes'
import { Ribbon, type FormatCmd, type SlidesViewMode } from './components/Ribbon'
import { contextElementTypeForNode, type ContextElementType } from './components/context-tabs'
import { SlideShowView } from './components/SlideShowView'
import { IconNotes, IconPlayBoxed } from './components/icons'
import { PresenterView } from './components/PresenterView'
import { CustomShowDialog } from './components/CustomShowDialog'
import { PrintDialog } from './components/PrintDialog'
import { FindReplaceDialog } from './components/FindReplaceDialog'
import { formatClock, type CustomShow } from './slideshow-utils'
import { ContextMenu } from './components/ContextMenu'
import { PasteOptionsFloater } from './components/PasteOptionsFloater'
import { FormatBackgroundPane, type BgPaneOp } from './components/FormatBackgroundPane'
import { FormatPane } from './components/FormatPane'
import { CommentsPane } from './components/CommentsPane'
import { AnimationPane } from './components/AnimationPane'
import { AnimPreviewOverlay } from './components/AnimatedSlide'
import { EquationDialog, HeaderFooterDialog, LinkDialog } from './components/InsertDialogs'
import { CutoutDialog } from './components/CutoutDialog'
import type { WordArtPreset } from '@genoffice/ui'
import type { ChartPresetDef, IconDef, SmartArtDef } from './insert-presets'
import { GensparkMark, IconAiBeautify, IconAiFactCheck, IconAiImage } from './components/icons'
import { ToastHost } from './components/toast'
import { showToast } from './components/toast-bus'
import { t, useI18n } from './i18n/locale'
import { AiPanel } from './ai/AiPanel'
import { ChartDataDialog } from './components/ChartDataDialog'
import type { BrushFormat } from './format-brush'
import { isTextUndoTarget, shouldRouteUndoToDeck } from './undo-routing'
import type {
  ActionCtx,
  CropTargetState,
  CtxMenuState,
  CutoutTargetState,
  EditingCellState,
  EditingState,
  HfDialogState,
  LinkDialogState,
  SlideShowState,
} from './action-context'
import { FIT_WIDTH } from './app-constants'
import { StageRuler } from './components/StageRuler'
import { formatRulerValue, type RulerUnit } from './ruler-ticks'
import * as fileActions from './file-actions'
import * as clipboardActions from './clipboard-actions'
import * as insertActions from './insert-actions'
import * as animationActions from './animation-actions'
import * as showActions from './show-actions'
import * as slideActions from './slide-actions'
import * as pictureEditActions from './picture-edit-actions'
import * as arrangeActions from './arrange-actions'
import * as tableActions from './table-actions'
import * as styleActions from './style-actions'
import { handleGlobalKeydown } from './keyboard-actions'
import { buildCtxItems } from './context-menu-items'

const _IS_MAC = navigator.platform.toLowerCase().includes('mac')

/** Effects the canvas can play as a one-shot click preview; 'random' resolves to one of these */
const PREVIEWABLE_TRANSITIONS: TransitionKind[] = [
  'morph',
  'fade',
  'push',
  'wipe',
  'split',
  'circle',
  'cover',
  'pull',
  'dissolve',
  'zoom',
]

/** Resizable thumbnail sidebar: drag the right edge; width persisted, clamped to a sane range */
const THUMBS_W_KEY = 'ai-slides-thumbs-width'
const THUMBS_W_DEFAULT = 120
const THUMBS_W_MIN = 110

function clampThumbsW(w: number): number {
  return Math.min(Math.max(w, THUMBS_W_MIN), Math.min(400, Math.round(window.innerWidth * 0.4)))
}

function loadThumbsW(): number {
  const saved = Number(localStorage.getItem(THUMBS_W_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampThumbsW(saved) : THUMBS_W_DEFAULT
}

/** Outline view: extract one page's text from the render tree (title = the text box with the largest font size, topmost). */
function outlineOf(s: RenderSlide): { title: string; lines: string[] } {
  const texts: Array<{ y: number; fontSize: number; text: string }> = []
  const walk = (nodes: readonly RenderNode[]) => {
    for (const n of nodes) {
      if ((n.type === 'shape' || n.type === 'text') && n.text?.lines?.length) {
        const text = n.text.lines
          .map((l) => l.runs.map((r) => r.text).join(''))
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim()
        if (text) {
          texts.push({
            y: n.box?.y ?? 0,
            fontSize: n.text.lines[0]?.runs[0]?.fontSizePx ?? 0,
            text,
          })
        }
      }
      if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
    }
  }
  walk(s.nodes)
  if (texts.length === 0) return { title: '', lines: [] }
  const title = [...texts].sort((a, b) => b.fontSize - a.fontSize || a.y - b.y)[0]!
  const lines = texts
    .filter((t) => t !== title)
    .sort((a, b) => a.y - b.y)
    .map((t) => t.text)
  return { title: title.text, lines }
}

/** Collect fonts/sizes (pt) of all text runs in a node (including group children, table cells) for ribbon display.
 * fontSizePx includes viewport scale and autofit fontScale; divide them out to get pt. */
function collectFontRuns(
  node: RenderNode,
  scale: number,
  out: Array<{ family: string; sizePt: number }>,
) {
  const pushLayout = (text?: {
    lines: Array<{
      runs: Array<{ fontFamily: string; srcFontFamily?: string; fontSizePx: number }>
    }>
    fontScale?: number
  }) => {
    const norm = scale * (text?.fontScale ?? 1) || 1
    for (const line of text?.lines ?? [])
      for (const run of line.runs)
        out.push({
          // Prefer the model's font name: fontFamily may be a missing-font substitution product
          // (DengXian on mac shows Arial otherwise, even though the file says DengXian)
          family: run.srcFontFamily ?? run.fontFamily,
          sizePt: Math.round((((run.fontSizePx / norm) * 72) / 96) * 2) / 2,
        })
  }
  if (node.type === 'shape' || node.type === 'text') pushLayout(node.text)
  else if (node.type === 'table') for (const cell of node.cells) pushLayout(cell.text)
  else if (node.type === 'group')
    for (const child of node.children) collectFontRuns(child, scale, out)
}

/** Per-paragraph bullet chars of one laid-out text body for the ribbon bullet gallery: '' for a
 * paragraph with no bullet, '#num' for numbered (matches no preset tile). Lines group into
 * paragraphs on paraStart so wrap continuations don't count. */
function collectBodyBulletChars(
  text:
    | { lines: Array<{ runs: Array<{ text: string; isBullet?: boolean }>; paraStart?: boolean }> }
    | undefined,
  out: Set<string>,
) {
  if (!text) return
  if (!text.lines.length) {
    out.add('')
    return
  }
  for (let i = 0; i < text.lines.length;) {
    let j = i + 1
    while (j < text.lines.length && !text.lines[j]!.paraStart) j++
    const bullet = text.lines
      .slice(i, j)
      .flatMap((l) => l.runs)
      .find((r) => r.isBullet)
    out.add(bullet ? (/^\d/.test(bullet.text) ? '#num' : bullet.text.trim()) : '')
    i = j
  }
}

/** Same collection across a node's text bodies (group children, all table cells). */
function collectBulletChars(node: RenderNode, out: Set<string>) {
  if (node.type === 'shape' || node.type === 'text') collectBodyBulletChars(node.text, out)
  else if (node.type === 'table')
    for (const cell of node.cells) collectBodyBulletChars(cell.text, out)
  else if (node.type === 'group') for (const child of node.children) collectBulletChars(child, out)
}

type ParaAlign = 'left' | 'center' | 'right' | 'justify'

/** Per-paragraph alignment of one laid-out text body: layout stamps `align` on every line
 * of a paragraph only when explicit, so an unset line reads as 'left' (the engine default —
 * some alignment is always current). Only paragraph-start lines count (wrap continuations
 * repeat the same value). */
function collectBodyAligns(
  text: { lines: Array<{ align?: ParaAlign; paraStart?: boolean }> } | undefined,
  out: Set<ParaAlign>,
) {
  if (!text) return
  if (!text.lines.length) {
    out.add('left')
    return
  }
  for (const line of text.lines) if (line.paraStart) out.add(line.align ?? 'left')
}

/** Same collection across a node's text bodies (group children, all table cells). */
function collectAligns(node: RenderNode, out: Set<ParaAlign>) {
  if (node.type === 'shape' || node.type === 'text') collectBodyAligns(node.text, out)
  else if (node.type === 'table') for (const cell of node.cells) collectBodyAligns(cell.text, out)
  else if (node.type === 'group') for (const child of node.children) collectAligns(child, out)
}

export function App() {
  const { lang } = useI18n()
  const [slides, setSlides] = useState<RenderSlide[]>([])
  // Layouts may reference Office-private fonts (resolved in main); register them as FontFaces
  useEffect(() => {
    if (slides.length) void syncPrivateFonts()
  }, [slides])
  const [path, setPath] = useState<string | null>(null)
  /** AiPanel reset key: incremented only on applyOpen (open/new file), not on draft path updates */
  const [aiPanelKey, setAiPanelKey] = useState(0)
  /** Theme body default font (fallback for the font box when the selection has no text element) */
  const [defaultFont, setDefaultFont] = useState<string | null>(null)
  const [current, setCurrent] = useState(0)
  const [selectedIds, setSelectedIds] = useState<string[]>([])
  /** Group being edited from inside (double-click to enter, click outside/Esc to exit); the selection may contain its children */
  const [enteredGroupId, setEnteredGroupId] = useState<string | null>(null)
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [editingCell, setEditingCell] = useState<EditingCellState | null>(null)
  // Armed shape draw mode (ribbon gallery pick); null = normal selection behavior
  const [drawKind, setDrawKind] = useState<InsertKind | null>(null)
  /** Latest-state bundle for the extracted action modules; refreshed every render (see action-context.ts). */
  const ctxRef = useRef<ActionCtx>(null as unknown as ActionCtx)
  const [zoom, setZoom] = useState(1)
  /** unscaled layout size of .stage-scale — its transform-scaled visual size is
   * scaleBox * zoom, which the wrapper zoom-box adopts so scrolling can reach it all */
  const stageScaleRef = useRef<HTMLDivElement | null>(null)
  const stageRelRef = useRef<HTMLDivElement | null>(null)
  const zoomBoxRef = useRef<HTMLDivElement | null>(null)
  const [scaleBox, setScaleBox] = useState<{ w: number; h: number } | null>(null)
  const [dirty, setDirty] = useState(false)
  const [status, setStatus] = useState('')
  // Status messages auto-dismiss after 4s: operation feedback needs only a brief showing; persistent info (page number/file name) lives on the left and in the title bar
  useEffect(() => {
    if (!status) return
    const t = window.setTimeout(() => setStatus(''), 4000)
    return () => window.clearTimeout(t)
  }, [status])
  const [showThumbs, setShowThumbs] = useState(true)
  // ── Thumbnail sidebar width (drag the divider to resize; persisted) ─────────
  const [thumbsW, setThumbsW] = useState(loadThumbsW)
  const thumbsListRef = useRef<HTMLDivElement | null>(null)
  // Re-clamp when the window shrinks (max is 40% of the window), like the AI panel
  useEffect(() => {
    const onResize = () => setThumbsW((w) => clampThumbsW(w))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])
  /** Drag to resize: width state follows the pointer (rAF-throttled so the Konva
   * thumbnails re-render at most once per frame); persisted on release. */
  const startThumbsResize = (e: React.PointerEvent<HTMLDivElement>) => {
    e.preventDefault()
    const list = thumbsListRef.current
    if (!list) return
    const left = list.getBoundingClientRect().left
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    let w = thumbsW
    let raf = 0
    const onMove = (ev: PointerEvent) => {
      w = clampThumbsW(ev.clientX - left)
      if (!raf)
        raf = requestAnimationFrame(() => {
          raf = 0
          setThumbsW(w)
        })
    }
    const onUp = () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      if (raf) cancelAnimationFrame(raf)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setThumbsW(w)
      localStorage.setItem(THUMBS_W_KEY, String(Math.round(w)))
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }
  const [autoSave, setAutoSave] = useState(
    () => localStorage.getItem('ai-slides-auto-save') === '1',
  )
  useEffect(() => {
    localStorage.setItem('ai-slides-auto-save', autoSave ? '1' : '0')
    window.slidesApi.setAutoSavePref?.(autoSave)
  }, [autoSave])
  const [showAi, setShowAi] = useState(() => localStorage.getItem('ai-slides-show-ai') !== '0')
  const [showFormat, setShowFormat] = useState(false)
  const [showBgFormat, setShowBgFormat] = useState(false)
  const [aiSettings, setAiSettings] = useState<AiSettings | null>(null)
  const [aiPreset, setAiPreset] = useState<{
    text: string
    nonce: number
    autoRun?: boolean
    displayText?: string
    attachments?: AttachmentMeta[]
    slideShot?: boolean
  } | null>(null)
  const [_recent, setRecent] = useState<string[]>([])
  const consumePendingRef = useRef<ReturnType<typeof window.slidesApi.consumePendingOpen> | null>(
    null,
  )
  const bootHandledRef = useRef(false)
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  const imageLoaderRef = useRef<ReturnType<typeof createImageLoader> | null>(null)
  const [hasClipboard, setHasClipboard] = useState(false)
  const [transition, setTransition] = useState<TransitionKind>('none')
  // ── Animations tab: current page's animation list + pane/preview ─────────────
  const [animations, setAnimations] = useState<AnimationItem[]>([])
  const [showAnimPane, setShowAnimPane] = useState(false)
  /** Animation pane selection (the timing controls bind to it first); -1 = none */
  const [selAnim, setSelAnim] = useState(-1)
  /** Canvas animation preview (plays when nonce>0; re-entering +1) */
  const [animPreview, setAnimPreview] = useState(0)
  /** Ribbon hover preview: the hovered effect played on the selection without applying */
  const [hoverAnim, setHoverAnim] = useState<{ nonce: number; items: AnimationItem[] } | null>(null)
  const [inkTool, setInkTool] = useState<InkTool>('select')
  const [inkPen, setInkPen] = useState<InkPenSettings>({ color: 'C00000', width: 2 })
  const [inkHighlighter, setInkHighlighter] = useState<InkPenSettings>({
    color: 'FFFF00',
    width: 10,
  })
  const [ctxMenu, setCtxMenu] = useState<CtxMenuState>(null)
  // ── Sections: grouping data + collapsed state + renaming state ────────────────
  const [sections, setSections] = useState<SectionInfo[]>([])
  const [collapsedSecs, setCollapsedSecs] = useState<Set<string>>(new Set())
  const [renamingSec, setRenamingSec] = useState<{ id: string; value: string } | null>(null)
  const stageWrapRef = useRef<HTMLDivElement | null>(null)
  const [stageViewportSize, setStageViewportSize] = useState({ w: 0, h: 0 })
  // ── View tab: view mode + display toggles ─────────────────────────────
  const [viewMode, setViewMode] = useState<SlidesViewMode>('normal')
  // Follow slide changes made outside the list (arrow keys, canvas paging); viewMode/showThumbs
  // re-run it because the list remounts at scroll 0 when the normal view returns
  useEffect(() => {
    thumbsListRef.current?.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' })
  }, [current, showThumbs, viewMode])
  const [masterItems, setMasterItems] = useState<MasterPartItem[] | null>(null)
  // ── Slide show: when non-null, covers the whole window (startAt is the start page index;
  //    customOrder = custom show playback sequence; rehearse = rehearsal timing mode) ────────
  const [slideShow, setSlideShow] = useState<SlideShowState | null>(null)
  /** Presenter view (single-window version; mutually exclusive with slideShow) */
  const [presenter, setPresenter] = useState<{ startAt: number } | null>(null)
  // ── Custom shows: document-level list (persisted to localStorage by file path) + management dialog ──
  const [customShows, setCustomShows] = useState<CustomShow[]>([])
  const [customShowDlgOpen, setCustomShowDlgOpen] = useState(false)
  const [findOpen, setFindOpen] = useState(false)
  /** Per-page dwell seconds awaiting confirmation after rehearsal (non-null shows the "save?" confirmation dialog) */
  const [pendingRehearse, setPendingRehearse] = useState<number[] | null>(null)
  const [showRuler, setShowRuler] = useState(false)
  const [showGrid, setShowGrid] = useState(false)
  const [showGuides, setShowGuides] = useState(false)
  // Draggable guides (pos = 0..1 relative to page width/height); persisted to localStorage by document path
  const [guides, setGuides] = useState<Array<{ axis: 'v' | 'h'; pos: number }>>([
    { axis: 'v', pos: 0.5 },
    { axis: 'h', pos: 0.5 },
  ])
  // PowerPoint model: the notes pane is SHOWN by default as a thin one-line bar
  // (type in place); the Notes buttons hide/show it entirely; drag its top edge
  // to any height, all the way down to hide.
  const [showNotes, setShowNotes] = useState(true)
  const [notesText, setNotesText] = useState('')
  /** Unsaved notes draft (flushed before page switch/save) */
  const notesDraftRef = useRef<{ index: number; text: string } | null>(null)
  /** Notes pane height (px): default = the thin one-line bar, drag-resizable */
  const [notesHeight, setNotesHeight] = useState(30)
  const notesDragRef = useRef<{ startY: number; startH: number } | null>(null)
  /** Splitter drag shared by the pane's top handle (startH = current height) and
   * the invisible pull-zone when hidden (startH = 0): below 20px = hidden
   * (restoring the pre-drag height for the Notes button), else live-resize. */
  const startNotesDrag = useCallback((e: React.MouseEvent, startH: number) => {
    e.preventDefault()
    notesDragRef.current = { startY: e.clientY, startH }
    const onMove = (ev: MouseEvent) => {
      const d = notesDragRef.current
      if (!d) return
      const raw = d.startH - (ev.clientY - d.startY)
      if (raw < 20) {
        setShowNotes(false)
        if (d.startH >= 30) setNotesHeight(Math.max(30, Math.min(480, d.startH)))
      } else {
        setShowNotes(true)
        setNotesHeight(Math.max(30, Math.min(480, raw)))
      }
    }
    const onUp = () => {
      notesDragRef.current = null
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
  }, [])
  // ── Review tab: comments ────────────────────────────────────────────────
  const [comments, setComments] = useState<SlideComment[]>([])
  const [showComments, setShowComments] = useState(false)
  const [commentsFocusNonce, setCommentsFocusNonce] = useState(0)
  /** Notes/comments must be re-fetched after undo/redo (they're not in RenderSlide) */
  const [annotationsNonce, setAnnotationsNonce] = useState(0)
  // ── Insert tab extensions: dialogs + screen recording ──────────────────────
  const [linkDialog, setLinkDialog] = useState<LinkDialogState | null>(null)
  const [hfDialog, setHfDialog] = useState<HfDialogState | null>(null)
  const [eqDialogOpen, setEqDialogOpen] = useState(false)
  const recorderRef = useRef<{ rec: MediaRecorder; stream: MediaStream } | null>(null)
  const [recording, setRecording] = useState(false)
  // ── Layout picking: layout list + slide size (loaded after the file opens) ─────────────────
  const [layoutsResult, setLayoutsResult] = useState<GetLayoutsResult | null>(null)
  // ── Picture crop mode ─────────────────────────────────────────────────────
  /** Non-null enters crop mode */
  const [cropTarget, setCropTarget] = useState<CropTargetState | null>(null)
  // ── Picture cutout (background removal) mode ───────────────────────────────
  /** Non-null opens the cutout dialog (dataUrl is the full original image data) */
  const [cutoutTarget, setCutoutTarget] = useState<CutoutTargetState | null>(null)
  // ── Format painter ───────────────────────────────────────────────────────
  /** Copied format (null = nothing copied yet) */
  const [brushFormat, setBrushFormat] = useState<BrushFormat | null>(null)
  /**
   * Format painter mode:
   *  - null    = not in format painter mode
   *  - 'once'  = paste format once (exits after one click)
   *  - 'continuous' = continuous mode (double-click the button to enter, Esc to exit)
   */
  const [brushMode, setBrushMode] = useState<'once' | 'continuous' | null>(null)

  const slide = slides[current]
  const hasDoc = !!slide

  /// True when no slide carries real content (only master decorations and
  /// untouched empty placeholders) — the ribbon's one-click AI actions grey out.
  const deckEmpty = useMemo(() => {
    const nodesHaveContent = (nodes: RenderNode[]): boolean =>
      nodes.some((n) => {
        if (n.decoration) return false
        if (n.type === 'group') return nodesHaveContent(n.children)
        if (n.type === 'picture' || n.type === 'table' || n.type === 'chart') return true
        if (n.type === 'shape' || n.type === 'text') {
          const hasText = (n.text?.lines ?? []).some((line) =>
            line.runs.some((run) => run.text.trim() !== ''),
          )
          // A drawn shape counts even without text; an untouched placeholder doesn't.
          return hasText || (n.type === 'shape' && !n.placeholder)
        }
        return false
      })
    return slides.every((s) => !nodesHaveContent(s.nodes))
  }, [slides])

  /** Write the notes draft back to the main process (called on page switch / blur / before save). */
  const flushNotes = useCallback(async () => {
    const pending = notesDraftRef.current
    if (!pending) return
    notesDraftRef.current = null
    const ok = await window.slidesApi.setNotes({ slideIndex: pending.index, text: pending.text })
    if (ok) setDirty(true)
  }, [])

  // Uncapped proportional fit ratio from the stage container's measured size
  // (fall back to a window estimate before mount)
  const rawFit = useCallback(
    (s?: RenderSlide) => {
      if (!s) return 1
      const el = stageWrapRef.current
      const viewportW =
        el?.clientWidth ||
        stageViewportSize.w ||
        window.innerWidth - (showThumbs ? thumbsW : 0) - (showAi ? 360 : 34)
      const viewportH = el?.clientHeight || stageViewportSize.h || window.innerHeight - 150
      const availW = viewportW - 56
      // -72: vertical padding is 48 (AI-bar headroom) + 32, minus the same 8px slack as width
      const availH = viewportH - 72
      return Math.min(availW / s.widthPx, availH / s.heightPx)
    },
    [showThumbs, showAi, stageViewportSize.h, stageViewportSize.w, thumbsW],
  )
  /** Fit zoom for display: rawFit clamped to the 0.1–1.5 auto-fit range */
  const fitZoom = useCallback(
    (s?: RenderSlide) => Math.min(1.5, Math.max(0.1, rawFit(s))),
    [rawFit],
  )
  // The Konva stage includes a transparent bleed around the slide for editing
  // off-page objects. That bleed must not create scrollbars while the slide
  // itself still fits; once the nominal slide overflows, normal panning resumes.
  const stageFitsViewport = !slide || zoom <= rawFit(slide) + 0.001
  /** Fit-pending flag after open: consumed when the editor's first frame mounts (stage container measurable) */
  const needsFitRef = useRef(false)
  /** Last auto-fit value: if current zoom still equals it → treated as "fit mode", re-fit on size changes */
  const lastFitRef = useRef<number | null>(null)
  const zoomLiveRef = useRef(1)
  useEffect(() => {
    zoomLiveRef.current = zoom
  }, [zoom])
  const slideLiveRef = useRef<RenderSlide | undefined>(undefined)
  useEffect(() => {
    slideLiveRef.current = slide
  }, [slide])
  // Live rawFit for the zoom-gesture code (previewZoom is dependency-free)
  const rawFitRef = useRef(rawFit)
  useEffect(() => {
    rawFitRef.current = rawFit
  }, [rawFit])

  /** While the slide fits the viewport, the centered position is the only valid one —
   * overflow is hidden, so the user cannot correct a stray scroll offset, and the
   * Konva bleed keeps scrollHeight above clientHeight so the browser never clamps
   * it back to 0 on its own. A zoom gesture's anchoring math can leave such an
   * offset behind (large Ctrl+wheel steps on Windows made the slide sit clipped
   * under the ribbon); zero it whenever the target zoom is at/below fit. */
  const resetFitScroll = useCallback((z: number) => {
    if (z > rawFitRef.current(slideLiveRef.current) + 0.001) return
    const wrap = stageWrapRef.current
    if (wrap && (wrap.scrollLeft !== 0 || wrap.scrollTop !== 0)) {
      wrap.scrollLeft = 0
      wrap.scrollTop = 0
    }
  }, [])

  // Non-gesture zoom paths (zoom-to-fit button, open-time fit, resize re-fit) land here
  useLayoutEffect(() => {
    if (stageFitsViewport) resetFitScroll(zoom)
  }, [stageFitsViewport, zoom, resetFitScroll])

  useLayoutEffect(() => {
    if (!slide || !needsFitRef.current) return
    needsFitRef.current = false
    const z = fitZoom(slide)
    lastFitRef.current = z
    setZoom(z)
  }, [slide, fitZoom])

  // measure the stage content's unscaled layout size (offsetWidth ignores the
  // transform); slide-size changes are the only thing that alters it
  useLayoutEffect(() => {
    const el = stageScaleRef.current
    if (!el) {
      setScaleBox(null)
      return
    }
    setScaleBox({ w: el.offsetWidth, h: el.offsetHeight })
  }, [slide?.widthPx, slide?.heightPx])

  // ── Stage rulers (PowerPoint-style fixed chrome outside the zoomed canvas) ──
  /** ruler unit follows the UI language: imperial for English, metric elsewhere */
  const rulerUnit: RulerUnit = lang === 'en' ? 'in' : 'cm'
  /** where slide coordinate 0 falls, in px relative to the stage viewport
   * (= ruler-local px, since each ruler strip is flush with .stage-wrap) */
  const [rulerOrigin, setRulerOrigin] = useState<{ x: number; y: number } | null>(null)
  useLayoutEffect(() => {
    if (!showRuler) return
    const wrap = stageWrapRef.current
    const rel = stageRelRef.current
    if (!wrap || !rel) return
    let raf = 0
    const measure = () => {
      raf = 0
      const w = wrap.getBoundingClientRect()
      const r = rel.getBoundingClientRect()
      const next = { x: r.left - w.left, y: r.top - w.top }
      setRulerOrigin((prev) =>
        prev && Math.abs(prev.x - next.x) < 0.5 && Math.abs(prev.y - next.y) < 0.5 ? prev : next,
      )
    }
    const queue = () => {
      if (!raf) raf = requestAnimationFrame(measure)
    }
    measure()
    wrap.addEventListener('scroll', queue, { passive: true })
    return () => {
      wrap.removeEventListener('scroll', queue)
      if (raf) cancelAnimationFrame(raf)
    }
    // stageViewportSize covers window/pane resizes; scaleBox covers slide-size changes
  }, [showRuler, zoom, stageViewportSize, scaleBox, viewMode, hasDoc])
  /** live guide being pulled off a ruler (slide fraction along its axis) */
  const [guidePreview, setGuidePreview] = useState<{ axis: 'v' | 'h'; pos: number } | null>(null)
  /** coordinate readout bubble during any guide drag (fixed viewport coords) */
  const [guideBubble, setGuideBubble] = useState<{ x: number; y: number; text: string } | null>(
    null,
  )
  /** pointer → slide fraction on the guide's cross axis ('h' guide ↔ y, 'v' guide ↔ x), unclamped */
  const guideFracAt = useCallback((axis: 'v' | 'h', client: { x: number; y: number }) => {
    const rect = stageRelRef.current?.getBoundingClientRect()
    if (!rect) return 0
    return axis === 'v' ? (client.x - rect.left) / rect.width : (client.y - rect.top) / rect.height
  }, [])

  // On container size changes (window/sidebar/thumbnail toggles): follow with a
  // re-fit while in fit mode, and clamp any manual zoom back down to fit whenever
  // the container gets too small — the canvas must never overflow the pane. A
  // manual zoom smaller than fit is left alone.
  useEffect(() => {
    const el = stageWrapRef.current
    if (!hasDoc || !el) return
    // Border-box size at the previous callback. Zooming past fit makes classic
    // scrollbars appear, which shrinks the observed content box without touching
    // the border box; treating that as "container got smaller" snapped every
    // zoom-in above fit straight back to fit. Only a border-box change is a real
    // window/pane resize allowed to clamp a manual zoom.
    let outerBox: { w: number; h: number } | null = null
    const ro = new ResizeObserver(() => {
      const nextSize = { w: el.clientWidth, h: el.clientHeight }
      setStageViewportSize((previous) =>
        previous.w === nextSize.w && previous.h === nextSize.h ? previous : nextSize,
      )
      const prevOuter = outerBox
      outerBox = { w: el.offsetWidth, h: el.offsetHeight }
      const z = fitZoom(slideLiveRef.current)
      const lf = lastFitRef.current
      const inFitMode = lf != null && Math.abs(zoomLiveRef.current - lf) <= 0.001
      if (!inFitMode) {
        // The overflow test uses the uncapped ratio: on large windows a manual
        // zoom above the 1.5 fit cap can still fit and must not be wiped
        if (zoomLiveRef.current <= rawFit(slideLiveRef.current) + 0.001) return
        const outerResized =
          prevOuter != null && (outerBox.w !== prevOuter.w || outerBox.h !== prevOuter.h)
        if (!outerResized) return
      }
      lastFitRef.current = z
      setZoom(z)
    })
    ro.observe(el)
    return () => ro.disconnect()
    // viewMode: reading/sorter unmount the editor, so the observer must re-bind
    // to the remounted .stage-wrap or window resizes stop re-fitting the canvas
  }, [hasDoc, fitZoom, rawFit, viewMode])

  const applyOpen = useCallback(
    (result: { path: string; slides: RenderSlide[]; defaultFont?: string } | null) => {
      if (!result) return
      setSlides(result.slides)
      setDefaultFont(result.defaultFont ?? null)
      setPath(result.path)
      setCurrent(0)
      setSelectedIds([])
      setEditing(null)
      setDirty(false)
      setInkTool('select')
      setAiPanelKey((k) => k + 1)
      needsFitRef.current = true
      setStatus(
        result.path
          ? t('appStatusOpened', {
              name: result.path.split('/').pop()!,
              count: result.slides.length,
            })
          : t('appStatusNewBlank'),
      )
      // Fetch the layout list asynchronously (doesn't block opening)
      void window.slidesApi.getLayouts().then((r) => setLayoutsResult(r))
    },
    [fitZoom],
  )

  const setSelectedId = useCallback((id: string | null, additive = false) => {
    setSelectedIds((prev) => {
      if (id == null) return []
      if (additive) return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
      return [id]
    })
  }, [])

  /** Group node being edited from inside (the group may no longer exist after page switch/undo; the effect exits as a fallback) */
  const enteredGroupNode = useMemo(() => {
    if (!enteredGroupId) return null
    const n = slides[current]?.nodes.find((x) => x.sourceId === enteredGroupId)
    return n?.type === 'group' ? (n as GroupRenderNode) : null
  }, [enteredGroupId, slides, current])
  useEffect(() => {
    if (enteredGroupId && !enteredGroupNode) setEnteredGroupId(null)
  }, [enteredGroupId, enteredGroupNode])

  /** Find a node by id (top level, or the group's children during in-group editing — the latter carries groupId for geometry commits) */
  const findNodeCtx = useCallback(
    (id: string): { node: RenderNode; groupId?: string } | null => {
      const top = slides[current]?.nodes.find((n) => n.sourceId === id)
      if (top) return { node: top }
      const c = enteredGroupNode?.children.find((x) => x.sourceId === id)
      return c ? { node: c, groupId: enteredGroupNode!.sourceId } : null
    },
    [slides, current, enteredGroupNode],
  )

  /** When id is a child of the group being edited, provide groupId (format ops go through the in-group patch pipeline) */
  const groupIdOf = useCallback(
    (id: string): string | undefined =>
      enteredGroupNode?.children.some((c) => c.sourceId === id)
        ? enteredGroupNode!.sourceId
        : undefined,
    [enteredGroupNode],
  )

  const openDialog = useCallback(async () => {
    const r = await window.slidesApi.openPptx(FIT_WIDTH)
    applyOpen(r)
  }, [applyOpen])

  // Save/export flows live in file-actions.ts; the editing-active flag lets ⌘S wait for the edit overlay to commit
  const editingActiveRef = useRef(false)
  editingActiveRef.current = !!editing || !!editingCell

  const save = useCallback(
    (quiet = false): Promise<boolean> => fileActions.save(ctxRef.current, quiet),
    [],
  )

  // Close guard (closing tab/window) chose "Save": run the full save flow and report the result
  useEffect(() => {
    return window.slidesApi.onCloseSaveRequest?.(() => {
      void save().then(
        (ok) => window.slidesApi.reportCloseSaveResult(ok),
        () => window.slidesApi.reportCloseSaveResult(false),
      )
    })
  }, [save])

  // Undo/redo stack occupancy pushed by the main process: the QAT buttons grey out when empty
  const [histState, setHistState] = useState({ canUndo: false, canRedo: false })
  useEffect(() => window.slidesApi.onHistoryChanged?.(setHistState), [])

  // Auto-save (off by default): when on and a file path exists, silently write back every 30s + on blur.
  // Saving rebuilds element ids; skipped during text editing/mouse-down (dragging) to avoid interrupting the current operation.
  const mouseDownRef = useRef(false)
  useEffect(() => {
    const down = () => {
      mouseDownRef.current = true
    }
    const up = () => {
      mouseDownRef.current = false
    }
    window.addEventListener('mousedown', down)
    window.addEventListener('mouseup', up)
    return () => {
      window.removeEventListener('mousedown', down)
      window.removeEventListener('mouseup', up)
    }
  }, [])
  useEffect(() => {
    if (!autoSave || !path) return
    let saving = false
    const tick = () => {
      if (saving || editing || editingCell || mouseDownRef.current) return
      void window.slidesApi.isDirty().then((d) => {
        if (!d || saving) return
        saving = true
        void save(true).finally(() => {
          saving = false
        })
      })
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [autoSave, path, editing, editingCell, save])

  const saveAs = useCallback(() => fileActions.saveAs(ctxRef.current), [])
  const exportImages = useCallback(() => fileActions.exportImages(ctxRef.current), [])
  const exportPdf = useCallback(() => fileActions.exportPdf(ctxRef.current), [])

  const [printDlgOpen, setPrintDlgOpen] = useState(false)

  /** Whether focus is in a text input (input/textarea/contentEditable) — these cases use native undo/delete */
  const inTextField = () => {
    const el = document.activeElement as HTMLElement | null
    return !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)
  }

  // Text dragged in plain DOM (e.g. AI panel) focuses body; canvas shape selection stays collapsed
  const hasDomTextSelection = () => {
    const sel = window.getSelection()
    return !!sel && !sel.isCollapsed
  }

  /** Apply the full slides set after undo/redo: page count may change (undoing a new page), clamp current */
  const applyHistoryResult = useCallback((r: RenderSlide[] | null) => {
    if (!r) return
    setSlides(r)
    setCurrent((c) => Math.min(c, r.length - 1))
    setSelectedIds([])
    setEditing(null)
    setPasteFloater(null) // The paste the floater refers to may have just been undone
    notesDraftRef.current = null // Undo overrides the unsaved draft, avoiding writing an old draft back
    setAnnotationsNonce((n) => n + 1) // Notes/comments aren't in RenderSlide; re-fetch
    void window.slidesApi.isDirty().then(setDirty)
  }, [])

  const undo = useCallback(async () => {
    // Preserve native undo while typing. The cleared AI composer explicitly yields to deck undo.
    const target = document.activeElement as HTMLElement | null
    if (editing || (isTextUndoTarget(target) && !shouldRouteUndoToDeck(target))) {
      document.execCommand('undo')
      return
    }
    applyHistoryResult(await window.slidesApi.undo())
  }, [editing, applyHistoryResult])

  const redo = useCallback(async () => {
    if (editing || inTextField()) {
      document.execCommand('redo')
      return
    }
    applyHistoryResult(await window.slidesApi.redo())
  }, [editing, applyHistoryResult])

  // Global shortcuts (keyboard-actions.ts): the handler reads the latest state via ctxRef, so attach once
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => handleGlobalKeydown(ctxRef.current, e)
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  // Zoom preview shared by trackpad pinch and the status-bar slider: a setZoom per event
  // re-renders the whole App (thumbnails included) dozens of times per gesture and froze
  // the pinch for seconds; instead the gesture only touches the CSS transform (rAF-batched)
  // and the state commit is debounced to the gesture's end.
  const zoomGestureRef = useRef<{
    raf: number
    timer: number
    pending: number | null
    anchor: { x: number; y: number } | null
  }>({
    raf: 0,
    timer: 0,
    pending: null,
    anchor: null,
  })
  /** `next` may be a functional updater (based on the pending gesture value, so rapid
   * steps compound correctly). `anchor` (client coords) is the screen point the zoom
   * pivots around — the cursor for wheel/pinch; defaults to the viewport center. */
  const previewZoom = useCallback(
    (next: number | ((current: number) => number), anchor?: { x: number; y: number }) => {
      const g = zoomGestureRef.current
      const target = typeof next === 'function' ? next(g.pending ?? zoomLiveRef.current) : next
      g.pending = Math.min(3, Math.max(0.25, target))
      g.anchor = anchor ?? null
      if (!g.raf) {
        g.raf = requestAnimationFrame(() => {
          g.raf = 0
          const scaleEl = stageScaleRef.current
          if (g.pending == null || !scaleEl) return
          const wrap = stageWrapRef.current
          // Content point (unscaled coords) currently under the anchor, measured from the
          // applied transform itself so it stays correct regardless of commit timing
          let ax = 0
          let ay = 0
          let px = 0
          let py = 0
          if (wrap) {
            const wr = wrap.getBoundingClientRect()
            ax = anchorClamp(g.anchor?.x, wr.left, wr.right) ?? wr.left + wr.width / 2
            ay = anchorClamp(g.anchor?.y, wr.top, wr.bottom) ?? wr.top + wr.height / 2
            const r = scaleEl.getBoundingClientRect()
            const applied = scaleEl.offsetWidth ? r.width / scaleEl.offsetWidth : 1
            px = (ax - r.left) / applied
            py = (ay - r.top) / applied
          }
          scaleEl.style.transform = `scale(${g.pending})`
          const box = zoomBoxRef.current
          if (box) {
            // offsetWidth ignores the transform = unscaled layout size (same basis as scaleBox)
            box.style.width = `${scaleEl.offsetWidth * g.pending}px`
            box.style.height = `${scaleEl.offsetHeight * g.pending}px`
          }
          // Scroll so the anchored content point stays under the anchor (the browser
          // clamps at the scroll edges; while the content still fits it stays centered)
          if (wrap) {
            const r = scaleEl.getBoundingClientRect()
            wrap.scrollLeft += px * g.pending - (ax - r.left)
            wrap.scrollTop += py * g.pending - (ay - r.top)
            // At/below fit the anchoring must not win over centering: the bleed's
            // scroll range would keep the offset alive after overflow turns hidden
            resetFitScroll(g.pending)
          }
          // Selection chrome counter-scales per frame so it holds a constant on-screen size
          window.dispatchEvent(new CustomEvent(ZOOM_PREVIEW_EVENT, { detail: g.pending }))
        })
      }
      window.clearTimeout(g.timer)
      g.timer = window.setTimeout(() => {
        if (g.pending == null) return
        setZoom(g.pending)
        // The commit may not change `zoom` (gesture ended where it started), so the
        // layout-effect reset would not re-run; clear any residue here as well
        resetFitScroll(g.pending)
        g.pending = null
        g.anchor = null
      }, 150)
    },
    [resetFitScroll],
  )
  useEffect(() => {
    const g = zoomGestureRef.current
    return () => {
      cancelAnimationFrame(g.raf)
      window.clearTimeout(g.timer)
    }
  }, [])

  // Trackpad pinch zoom: Chromium synthesizes pinch gestures as ctrlKey+wheel events;
  // Ctrl/⌘ + wheel zoom also supported. Needs passive:false to preventDefault.
  // Plain wheel/trackpad scroll turns pages (PowerPoint/WPS behavior) — but only
  // while the slide fits the viewport; zoomed-in overflow keeps native panning.
  useEffect(() => {
    const el = stageWrapRef.current
    if (!el) return
    const pager = createWheelPager()
    const onWheel = (ev: WheelEvent) => {
      if (ev.ctrlKey || ev.metaKey) {
        ev.preventDefault()
        const factor = Math.exp(-ev.deltaY * 0.01)
        previewZoom((z) => z * factor, { x: ev.clientX, y: ev.clientY })
        return
      }
      // Never flip out from under a live text edit — the overlay's commit
      // must not depend on an unmount blur (same guard as autosave/⌘S)
      if (editingActiveRef.current) return
      const fits = zoomLiveRef.current <= rawFitRef.current(slideLiveRef.current) + 0.001
      if (!fits) return
      const flip = pager.feed(ev.deltaY, ev.timeStamp)
      if (flip) setCurrent((c) => Math.min(Math.max(c + flip, 0), slides.length - 1))
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  }, [hasDoc, viewMode, previewZoom, slides.length])

  const newBlank = useCallback(async () => {
    const r = await window.slidesApi.newBlank(FIT_WIDTH)
    applyOpen(r)
    return r
  }, [applyOpen])

  // Files are pushed open by the main process (double-click/command line); with no pending file the
  // window lands directly in the editor on a fresh blank deck (the AI panel carries the generate-from-prompt flow).
  // StrictMode runs the mount effect twice, but the pending queue can only be consumed once, so the
  // consume Promise is stored in a shared ref and its result is processed only once.
  useEffect(() => {
    // newBlank itself can fail (IPC/main-process error); one retry for transient
    // hiccups, then surface the error — otherwise the tab silently sticks on
    // "Opening…" forever and the click that created it looks like a no-op
    const bootBlank = () =>
      newBlank().catch((err: unknown) => {
        void newBlank().catch(() => {
          showToast(err instanceof Error ? err.message : String(err), 'error')
        })
      })
    const off = window.slidesApi.onOpened((r) => applyOpen(r))
    consumePendingRef.current ??= window.slidesApi.consumePendingOpen(FIT_WIDTH)
    void consumePendingRef.current
      .then((r) => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        if (r) applyOpen(r)
        else void bootBlank()
      })
      // Open failures (corrupt file etc.) also land on a blank deck, or it stays at "Opening…" forever
      .catch(() => {
        if (bootHandledRef.current) return
        bootHandledRef.current = true
        void bootBlank()
      })
    return off
  }, [applyOpen, newBlank])

  // File renamed externally (shell Home list rename) → sync the title-bar path (content unchanged, dirty untouched)
  useEffect(() => window.slidesApi.onRenamed((p) => setPath(p)), [])

  useEffect(() => {
    void window.slidesApi.getAiSettings().then(setAiSettings)
  }, [])

  // Recent files for the start screen
  useEffect(() => {
    if (slides.length === 0) void window.slidesApi.getRecentFiles().then(setRecent)
  }, [slides.length])

  const toggleAi = useCallback(() => {
    setShowAi((v) => {
      localStorage.setItem('ai-slides-show-ai', v ? '0' : '1')
      return !v
    })
  }, [])

  const pushAiPreset = useCallback(
    (
      text: string,
      autoRun = true,
      displayText?: string,
      attachments?: AttachmentMeta[],
      slideShot?: boolean,
    ) => {
      setShowAi(() => {
        localStorage.setItem('ai-slides-show-ai', '1')
        return true
      })
      setAiPreset({
        text,
        nonce: Date.now(),
        autoRun,
        displayText,
        ...(attachments && attachments.length > 0 ? { attachments } : {}),
        ...(slideShot ? { slideShot } : {}),
      })
    },
    [],
  )

  const applySlide = useCallback((slideIndex: number, updated: RenderSlide) => {
    setSlides((s) => s.map((sl, i) => (i === slideIndex ? updated : sl)))
    setDirty(true)
  }, [])

  const insertElement = useCallback(
    (kind: InsertKind) => insertActions.insertElement(ctxRef.current, kind),
    [],
  )
  // Shape draw mode (PowerPoint/WPS parity): gallery pick arms the crosshair, the canvas commits the box.
  // The armed kind lives in a ref too: the commit side effect must stay out of the
  // setState updater (StrictMode double-invokes updaters → double insert).
  const drawKindRef = useRef<InsertKind | null>(null)
  const pickShape = useCallback((kind: InsertKind) => {
    setEditing(null)
    setEditingCell(null)
    setSelectedIds([])
    drawKindRef.current = kind
    setDrawKind(kind)
  }, [])
  const commitDraw = useCallback((rect: DrawRect) => {
    const kind = drawKindRef.current
    drawKindRef.current = null
    setDrawKind(null)
    if (kind) void insertActions.insertShapeAt(ctxRef.current, kind, rect)
  }, [])
  const insertImage = useCallback(() => insertActions.insertImage(ctxRef.current), [])

  const onBackground = useCallback((op: BgPaneOp, allSlides?: boolean) => {
    const cur = ctxRef.current.current
    return styleActions.onBackground(ctxRef.current, {
      ...op,
      slideIndex: allSlides ? -1 : cur,
      // Mode change / apply-to-all reuses the current slide's background image
      ...(op.kind === 'image' && op.pick === false ? { sourceSlideIndex: cur } : {}),
    })
  }, [])
  const applyThemePreset = useCallback(
    (preset: SlideThemePreset) => styleActions.applyThemePreset(ctxRef.current, preset),
    [],
  )
  const onStroke = useCallback(
    (sourceId: string, stroke: { color: string; widthPt: number; dash?: string } | null) =>
      styleActions.onStroke(ctxRef.current, sourceId, stroke),
    [],
  )

  const applyDeck = useCallback((all: RenderSlide[], goTo?: number) => {
    setSlides(all)
    if (goTo != null) setCurrent(goTo)
    setSelectedIds([])
    setEditing(null)
    setDirty(true)
  }, [])

  const addSlide = useCallback(() => slideActions.addSlide(ctxRef.current), [])
  const addSlideWithLayout = useCallback(
    (layoutPath: string) => slideActions.addSlideWithLayout(ctxRef.current, layoutPath),
    [],
  )

  const deleteSelected = useCallback(() => clipboardActions.deleteSelected(ctxRef.current), [])
  const copySelected = useCallback(() => clipboardActions.copySelected(ctxRef.current), [])
  const cutSelected = useCallback(() => clipboardActions.cutSelected(ctxRef.current), [])

  /** Insert an external image at page center at natural size (clamped to half the page). */
  const insertExternalImage = useCallback(
    (base64: string, ext: string, atPx?: { x: number; y: number }) =>
      clipboardActions.insertExternalImage(ctxRef.current, base64, ext, atPx),
    [],
  )

  // The clipboard lives in the main process, so a slide copied here can be pasted
  // into any other open deck.
  const [canPasteSlide, setCanPasteSlide] = useState(false)

  // Paste-options floater on the just-pasted slide's thumbnail (PowerPoint-style);
  // clicking an option undoes that paste and redoes it with the chosen mode
  const [pasteFloater, setPasteFloater] = useState<{
    index: number
    mode: PasteSlideMode
  } | null>(null)
  const repasteSlideAs = useCallback(
    (mode: PasteSlideMode) => void clipboardActions.repasteSlideAs(ctxRef.current, mode),
    [],
  )

  const pasteClipboard = useCallback(() => clipboardActions.pasteClipboard(ctxRef.current), [])
  const duplicateSelected = useCallback(
    (ids: string[], dxPx = 16, dyPx = 16) =>
      clipboardActions.duplicateSelected(ctxRef.current, ids, dxPx, dyPx),
    [],
  )

  // ── Format painter operations (clipboard-actions.ts) ────────────────────────
  const onFormatBrushClick = useCallback(
    () => clipboardActions.onFormatBrushClick(ctxRef.current),
    [],
  )
  const onFormatBrushDoubleClick = useCallback(
    () => clipboardActions.onFormatBrushDoubleClick(ctxRef.current),
    [],
  )
  const applyBrushToElement = useCallback(
    (targetId: string) => clipboardActions.applyBrushToElement(ctxRef.current, targetId),
    [],
  )

  /**
   * Format-painter-aware select handler: while the painter is active, intercept clicks →
   * applyBrushToElement, without changing the selection (existing selection preserved).
   */
  const handleCanvasSelect = useCallback(
    (id: string | null, additive = false) => {
      if (brushMode && id) {
        void applyBrushToElement(id)
        return
      }
      // Clicking blank/an element outside the group during in-group editing = exit the group
      if (
        enteredGroupId &&
        (id == null || !enteredGroupNode?.children.some((c) => c.sourceId === id))
      ) {
        setEnteredGroupId(null)
      }
      setSelectedId(id, additive)
    },
    [brushMode, applyBrushToElement, setSelectedId, enteredGroupId, enteredGroupNode],
  )

  const handleEnterGroup = useCallback((groupId: string, childId: string | null) => {
    setEnteredGroupId(groupId)
    setSelectedIds(childId ? [childId] : [])
  }, [])

  const insertTable = useCallback(
    (rows: number, cols: number) => insertActions.insertTable(ctxRef.current, rows, cols),
    [],
  )

  // ── Insert tab extensions (insert-actions.ts): icons / charts / SmartArt / WordArt / fields / links / Zoom / footer / equations / media ──
  const insertIcon = useCallback(
    (def: IconDef, color: string) => insertActions.insertIcon(ctxRef.current, def, color),
    [],
  )
  const insertChart = useCallback(
    (kind: ChartPresetDef['kind']) => insertActions.insertChart(ctxRef.current, kind),
    [],
  )
  const insertSmartArt = useCallback(
    (def: SmartArtDef) => insertActions.insertSmartArt(ctxRef.current, def),
    [],
  )
  const insertWordArt = useCallback(
    (preset: WordArtPreset) => insertActions.insertWordArt(ctxRef.current, preset),
    [],
  )
  const insertField = useCallback(
    (type: 'datetime' | 'slidenum') => insertActions.insertField(ctxRef.current, type),
    [],
  )
  const openLinkDialog = useCallback(() => insertActions.openLinkDialog(ctxRef.current), [])
  const applyLink = useCallback(
    (target: LinkTargetOp | null) => insertActions.applyLink(ctxRef.current, target),
    [],
  )
  const insertZoom = useCallback(
    (target: number) => insertActions.insertZoom(ctxRef.current, target),
    [],
  )
  const openHeaderFooter = useCallback(() => insertActions.openHeaderFooter(ctxRef.current), [])
  const applyHf = useCallback(
    (opts: { footer: string | null; slideNum: boolean; date: string | null; dateAuto: boolean }) =>
      insertActions.applyHf(ctxRef.current, opts),
    [],
  )
  const insertEquation = useCallback(
    (text: string) => insertActions.insertEquation(ctxRef.current, text),
    [],
  )
  const insertMediaFile = useCallback(
    (kind: 'video' | 'audio') => insertActions.insertMediaFile(ctxRef.current, kind),
    [],
  )
  const insertModel3dFile = useCallback(() => insertActions.insertModel3dFile(ctxRef.current), [])
  const toggleScreenRecord = useCallback(() => insertActions.toggleScreenRecord(ctxRef.current), [])

  // Reflect the current page's transition effect on page/document changes
  useEffect(() => {
    if (!hasDoc) return
    void window.slidesApi.getTransition(current).then(setTransition)
  }, [hasDoc, current, path])

  const applyTransition = useCallback(
    (kind: TransitionKind, allSlides: boolean) =>
      animationActions.applyTransition(ctxRef.current, kind, allSlides),
    [],
  )

  // ── Transitions tab: one-shot canvas preview when an effect is clicked (PPT-style) ──
  const [transPreviewKind, setTransPreviewKind] = useState<TransitionKind | null>(null)
  const previewTransitionOnCanvas = useCallback((kind: TransitionKind) => {
    const concrete =
      kind === 'random'
        ? PREVIEWABLE_TRANSITIONS[Math.floor(Math.random() * PREVIEWABLE_TRANSITIONS.length)]!
        : kind
    if (concrete === 'none') return
    // drop the class for one frame so re-clicking the same effect restarts its animation
    setTransPreviewKind(null)
    requestAnimationFrame(() => setTransPreviewKind(concrete))
  }, [])

  // ── Animations tab: current page's animation list (refreshed after page switch/edit/undo; re-fetched whenever the slide identity changes) ──
  useEffect(() => {
    if (!hasDoc) {
      setAnimations([])
      return
    }
    let cancelled = false
    void window.slidesApi.getAnimations(current).then((items) => {
      if (!cancelled) setAnimations(items)
    })
    return () => {
      cancelled = true
    }
  }, [hasDoc, current, path, slide, annotationsNonce])

  // Guides persisted per document (loaded on document switch, saved on change)
  useEffect(() => {
    if (!path) return
    try {
      const raw = localStorage.getItem(`ai-slides-guides:${path}`)
      setGuides(
        raw
          ? JSON.parse(raw)
          : [
              { axis: 'v', pos: 0.5 },
              { axis: 'h', pos: 0.5 },
            ],
      )
    } catch {
      setGuides([
        { axis: 'v', pos: 0.5 },
        { axis: 'h', pos: 0.5 },
      ])
    }
  }, [path])
  useEffect(() => {
    if (!path) return
    try {
      localStorage.setItem(`ai-slides-guides:${path}`, JSON.stringify(guides))
    } catch {
      /* Quota full: guides aren't critical data, ignore */
    }
  }, [path, guides])

  // Clear pane selection and preview on page switch
  useEffect(() => {
    setSelAnim(-1)
    setAnimPreview(0)
    setHoverAnim(null)
  }, [current, path])

  // Animation actions live in animation-actions.ts
  const hoverPreviewAnimation = useCallback(
    (effect: AnimEffectKind, motionPath?: string) =>
      animationActions.hoverPreviewAnimation(ctxRef.current, effect, motionPath),
    [],
  )

  // "By paragraph" toggle: when applying entrance animations, split into one animation per target paragraph (one click reveals one)
  const [animByParagraph, setAnimByParagraph] = useState(false)

  const applyAnimation = useCallback(
    (effect: AnimEffectKind | 'none') => animationActions.applyAnimation(ctxRef.current, effect),
    [],
  )
  const addAnimation = useCallback(
    (effect: AnimEffectKind) => animationActions.addAnimation(ctxRef.current, effect),
    [],
  )
  const applyMotionPath = useCallback(
    (path: string) => animationActions.applyMotionPath(ctxRef.current, path),
    [],
  )

  /** Index of the animation bound to the timing controls: pane selection first, else the selected shape's last one. */
  const timingIdx = useMemo(() => {
    if (selAnim >= 0 && selAnim < animations.length) return selAnim
    if (selectedIds.length > 0) {
      for (let i = animations.length - 1; i >= 0; i--) {
        if (animations[i]!.sourceId === selectedIds[0]) return i
      }
    }
    return -1
  }, [selAnim, animations, selectedIds])

  const patchAnimTiming = useCallback(
    (patch: { trigger?: AnimTrigger; durationMs?: number; delayMs?: number }) =>
      animationActions.patchAnimTiming(ctxRef.current, patch),
    [],
  )
  const moveAnimation = useCallback(
    (idx: number, dir: -1 | 1) => animationActions.moveAnimation(ctxRef.current, idx, dir),
    [],
  )
  const deleteAnimation = useCallback(
    (idx: number) => animationActions.deleteAnimation(ctxRef.current, idx),
    [],
  )

  /** Selected shape's current animation effect (gallery highlight: first match). */
  const selectedAnimEffect = useMemo(() => {
    if (selectedIds.length === 0) return null
    return animations.find((a) => a.sourceId === selectedIds[0])?.effect ?? null
  }, [selectedIds, animations])

  const toggleAnimPane = useCallback(() => {
    setShowAnimPane((v) => {
      if (!v) {
        setShowFormat(false)
        setShowBgFormat(false)
        setShowComments(false)
      }
      return !v
    })
  }, [])

  /** Open the format-background pane (Design tab button / canvas context menu) */
  const openBgFormat = useCallback(() => {
    setShowBgFormat(true)
    setShowFormat(false)
    setShowAnimPane(false)
    setShowComments(false)
  }, [])

  // ── Slide show tab (show-actions.ts): start show / presenter view / hide slide ──
  const startSlideShow = useCallback(
    (fromStart: boolean) => showActions.startSlideShow(ctxRef.current, fromStart),
    [],
  )
  const exitSlideShow = useCallback(
    (lastIndex: number) => showActions.exitSlideShow(ctxRef.current, lastIndex),
    [],
  )

  // ── Custom shows: load the document's list on document switch (no path = unsaved new document, memory-only) ──
  useEffect(() => {
    setCustomShowDlgOpen(false)
    if (!path) {
      setCustomShows([])
      return
    }
    try {
      const raw = localStorage.getItem(`ai-slides-custom-shows:${path}`)
      const parsed: unknown = raw ? JSON.parse(raw) : []
      setCustomShows(Array.isArray(parsed) ? (parsed as CustomShow[]) : [])
    } catch {
      setCustomShows([])
    }
  }, [path])

  const updateCustomShows = useCallback(
    (shows: CustomShow[]) => showActions.updateCustomShows(ctxRef.current, shows),
    [],
  )
  const playCustomShow = useCallback(
    (show: CustomShow) => showActions.playCustomShow(ctxRef.current, show),
    [],
  )
  const startRehearseShow = useCallback(() => showActions.startRehearseShow(ctxRef.current), [])
  const onRehearseDone = useCallback(
    (perPageSec: number[]) => showActions.onRehearseDone(ctxRef.current, perPageSec),
    [],
  )
  const saveRehearseTimings = useCallback(() => showActions.saveRehearseTimings(ctxRef.current), [])
  const startPresenterView = useCallback(
    (fromStart: boolean) => showActions.startPresenterView(ctxRef.current, fromStart),
    [],
  )
  const exitPresenterView = useCallback(
    (lastIndex: number) => showActions.exitPresenterView(ctxRef.current, lastIndex),
    [],
  )
  const switchPresenterToShow = useCallback(
    (lastIndex: number) => showActions.switchPresenterToShow(ctxRef.current, lastIndex),
    [],
  )
  const toggleHidden = useCallback(
    (index: number) => showActions.toggleHidden(ctxRef.current, index),
    [],
  )

  // Picture crop / cutout (picture-edit-actions.ts)
  const startCrop = useCallback(() => pictureEditActions.startCrop(ctxRef.current), [])
  const commitCrop = useCallback(
    (rect: { l: number; t: number; r: number; b: number } | null) =>
      pictureEditActions.commitCrop(ctxRef.current, rect),
    [],
  )
  const cancelCrop = useCallback(() => pictureEditActions.cancelCrop(ctxRef.current), [])
  const startCutout = useCallback(() => pictureEditActions.startCutout(ctxRef.current), [])
  const applyCutout = useCallback(
    (pngDataUrl: string) => pictureEditActions.applyCutout(ctxRef.current, pngDataUrl),
    [],
  )

  // Grouping / alignment (arrange-actions.ts)
  const alignSelected = useCallback(
    (op: Parameters<typeof arrangeActions.alignSelected>[1]) =>
      arrangeActions.alignSelected(ctxRef.current, op),
    [],
  )
  const flipSelected = useCallback(
    (axis: 'h' | 'v') => arrangeActions.flipSelected(ctxRef.current, axis),
    [],
  )

  // Loaded on every slide switch (not just while the pane is open): the
  // collapsed add-notes bar shows the current slide's first note line
  useEffect(() => {
    if (!hasDoc) return
    let cancelled = false
    void flushNotes()
      .then(() => window.slidesApi.getNotes(current))
      .then((t) => {
        if (!cancelled) setNotesText(t)
      })
    return () => {
      cancelled = true
    }
  }, [hasDoc, current, path, annotationsNonce, flushNotes])

  const onNotesChange = useCallback(
    (text: string) => {
      setNotesText(text)
      notesDraftRef.current = { index: current, text }
    },
    [current],
  )

  // ── Comments: fetch the current page's list on page switch/document change/undo (the ribbon badge uses it too) ──────────
  useEffect(() => {
    if (!hasDoc) {
      setComments([])
      return
    }
    let cancelled = false
    void window.slidesApi.getComments(current).then((c) => {
      if (!cancelled) setComments(c)
    })
    return () => {
      cancelled = true
    }
  }, [hasDoc, current, path, annotationsNonce])

  const addComment = useCallback(
    async (text: string) => {
      const r = await window.slidesApi.addComment({ slideIndex: current, text })
      if (r) {
        setComments(r)
        setDirty(true)
        setStatus(t('appStatusCommentAdded'))
      }
    },
    [current],
  )

  const deleteComment = useCallback(
    async (c: SlideComment) => {
      const r = await window.slidesApi.deleteComment({
        slideIndex: current,
        authorId: c.authorId,
        idx: c.idx,
      })
      if (r) {
        setComments(r)
        setDirty(true)
      }
    },
    [current],
  )

  const openComments = useCallback((focus: boolean) => {
    setShowComments(true)
    setShowFormat(false)
    setShowBgFormat(false)
    setShowAnimPane(false)
    if (focus) setCommentsFocusNonce((n) => n + 1)
  }, [])

  // ── View modes ──────────────────────────────────────────────────────────
  const onViewMode = useCallback((mode: SlidesViewMode) => {
    setViewMode(mode)
    setEditing(null)
    setSelectedIds([])
    setCtxMenu(null)
  }, [])

  // ── Master editing view (full-screen overlay; edits go through slides:master-* IPC) ────────────
  const enterMasterView = useCallback(async () => {
    if (!hasDoc) return
    setEditing(null)
    setEditingCell(null)
    setSelectedIds([])
    setCtxMenu(null)
    const r = await window.slidesApi.masterEnter(FIT_WIDTH)
    if (r?.items.length) setMasterItems(r.items)
  }, [hasDoc])

  const closeMasterView = useCallback(async () => {
    const all = await window.slidesApi.masterClose()
    setMasterItems(null)
    if (all) {
      // The inheritance chain changed: replace the whole render tree (all-new element ids)
      setSlides(all)
      setSelectedIds([])
    }
    void window.slidesApi.isDirty().then(setDirty)
  }, [])

  const zoomToFit = useCallback(() => {
    const z = fitZoom(slide)
    lastFitRef.current = z
    setZoom(z)
  }, [fitZoom, slide])

  // Reading view: ←/→/↑/↓/space page turn, Esc exit (capture beats the generic shortcuts)
  useEffect(() => {
    if (viewMode !== 'reading') return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        setViewMode('normal')
      } else if (
        e.key === 'ArrowRight' ||
        e.key === 'ArrowDown' ||
        e.key === ' ' ||
        e.key === 'PageDown' ||
        e.key === 'Enter'
      ) {
        e.preventDefault()
        setCurrent((c) => Math.min(c + 1, slides.length - 1))
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault()
        setCurrent((c) => Math.max(c - 1, 0))
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [viewMode, slides.length])

  // Reading view follows window size
  const [winSize, setWinSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  useEffect(() => {
    if (viewMode !== 'reading') return
    const onResize = () => setWinSize({ w: window.innerWidth, h: window.innerHeight })
    onResize()
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [viewMode])

  // ── Section management ─────────────────────────────────────────────────
  // Section data follows document refreshes: open/edit/undo/save all swap the slides array reference; re-fetch here in one place
  useEffect(() => {
    if (!hasDoc) {
      setSections([])
      return
    }
    let alive = true
    void window.slidesApi.getSections().then((r) => {
      if (alive) setSections(r ?? [])
    })
    return () => {
      alive = false
    }
  }, [hasDoc, slides])

  const addSectionAt = useCallback(
    (index: number) => slideActions.addSectionAt(ctxRef.current, index),
    [],
  )

  // ── Thumbnail drag reorder ────────────────────────────────────────────────
  const [dragThumb, setDragThumb] = useState<number | null>(null)
  /** Insert position: 0..slides.length, value k means insert before page k */
  const [dropPos, setDropPos] = useState<number | null>(null)

  const moveSlideTo = useCallback(
    (from: number, insertAt: number) => slideActions.moveSlideTo(ctxRef.current, from, insertAt),
    [],
  )

  /** Drag props shared by the thumbnail list / sorter view; horizontal = sorter grid (front/back half decided by X) */
  const thumbDragProps = (i: number, horizontal = false) => ({
    draggable: true,
    onDragStart: (e: React.DragEvent) => {
      e.dataTransfer.effectAllowed = 'move'
      setDragThumb(i)
    },
    onDragOver: (e: React.DragEvent) => {
      if (dragThumb == null) return
      e.preventDefault()
      e.dataTransfer.dropEffect = 'move'
      const r = e.currentTarget.getBoundingClientRect()
      const before = horizontal
        ? e.clientX < r.left + r.width / 2
        : e.clientY < r.top + r.height / 2
      setDropPos(before ? i : i + 1)
    },
    onDrop: (e: React.DragEvent) => {
      e.preventDefault()
      // Recompute at the drop point (don't read dropPos state: the last dragover's setState may not have committed yet)
      if (dragThumb != null) {
        const r = e.currentTarget.getBoundingClientRect()
        const before = horizontal
          ? e.clientX < r.left + r.width / 2
          : e.clientY < r.top + r.height / 2
        void moveSlideTo(dragThumb, before ? i : i + 1)
      }
      setDragThumb(null)
      setDropPos(null)
    },
    onDragEnd: () => {
      setDragThumb(null)
      setDropPos(null)
    },
  })

  /** Drag visual state classes: source page dragging; drop point k draws a line on page k's top edge, the end drop point on the last page's bottom edge */
  const thumbDragCls = (i: number) =>
    `${dragThumb === i ? ' dragging' : ''}${
      dropPos === i
        ? ' drop-before'
        : dropPos === i + 1 && i === slides.length - 1
          ? ' drop-after'
          : ''
    }`

  const commitRenameSection = useCallback(
    () => slideActions.commitRenameSection(ctxRef.current),
    [],
  )

  const toggleSection = useCallback((id: string) => {
    setCollapsedSecs((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }, [])

  /**
   * Sidebar grouping: split the page sequence by each section's first-page index (section i
   * covers [start_i, start_{i+1})); pages before the first section start go into an
   * "unsectioned" group — tolerating stale section data after page insertions/deletions.
   */
  const sectionGroups = useMemo(() => {
    if (!sections.length || !slides.length) return null
    const total = slides.length
    const starts = new Array<number>(sections.length)
    let nextStart = total
    for (let i = sections.length - 1; i >= 0; i--) {
      const own = sections[i]!.slideIndices.length
        ? Math.min(...sections[i]!.slideIndices)
        : nextStart
      starts[i] = Math.min(own, nextStart)
      nextStart = starts[i]!
    }
    const groups: Array<{ id: string | null; name: string; start: number; end: number }> = []
    if (starts[0]! > 0)
      groups.push({ id: null, name: t('appSectionDefault'), start: 0, end: starts[0]! })
    sections.forEach((s, i) => {
      groups.push({
        id: s.id,
        name: s.name,
        start: starts[i]!,
        end: i + 1 < sections.length ? starts[i + 1]! : total,
      })
    })
    return groups
  }, [sections, slides.length, lang])

  /** Canvas right-click: select the hit element first (replace the selection if it isn't in it), clear selection on blank */
  const onCanvasContextMenu = useCallback(
    (sourceId: string | null, x: number, y: number, cell?: { row: number; col: number }) => {
      if (editing) return
      if (sourceId) {
        setSelectedIds((prev) => (prev.includes(sourceId) ? prev : [sourceId]))
        setCtxMenu({ kind: 'element', x, y, targetId: sourceId, ...(cell ? { cell } : {}) })
      } else {
        setSelectedIds([])
        setCtxMenu({ kind: 'canvas', x, y })
      }
    },
    [editing],
  )

  // Menu commands. Cut/copy/paste dispatch by context: text mode goes back to the native clipboard, canvas mode uses the element clipboard
  useEffect(() => {
    return window.slidesApi.onMenuCommand((cmd) => {
      // Master view: only allow save (undo/clipboard etc. target normal pages, not applicable inside the master)
      if (masterItems) {
        if (cmd === 'save') void save()
        return
      }
      if (cmd === 'open') void openDialog()
      else if (cmd === 'save') void save()
      else if (cmd === 'save-as') void saveAs()
      // macOS has no File ribbon tab, so these only exist in the menu
      else if (cmd === 'export-pdf') void exportPdf()
      else if (cmd === 'export-images') void exportImages()
      else if (cmd === 'print') setPrintDlgOpen(true)
      // Through the preview path so the zoom pivots on the viewport center, not the scroll origin
      else if (cmd === 'zoom-in') previewZoom((z) => Math.min(z * 1.15, 3))
      else if (cmd === 'zoom-out') previewZoom((z) => Math.max(z / 1.15, 0.25))
      else if (cmd === 'zoom-reset') previewZoom(1)
      else if (cmd === 'undo') void undo()
      else if (cmd === 'redo') void redo()
      else if (cmd === 'cut') {
        if (editing || inTextField() || hasDomTextSelection())
          void window.slidesApi.nativeClipboard('cut')
        else void cutSelected()
      } else if (cmd === 'copy') {
        if (editing || inTextField() || hasDomTextSelection())
          void window.slidesApi.nativeClipboard('copy')
        else void copySelected()
      } else if (cmd === 'paste') {
        if (editing || inTextField()) void window.slidesApi.nativeClipboard('paste')
        else void pasteClipboard()
      }
    })
  }, [
    openDialog,
    save,
    saveAs,
    exportPdf,
    exportImages,
    undo,
    redo,
    editing,
    cutSelected,
    copySelected,
    pasteClipboard,
    masterItems,
    previewZoom,
  ])

  // Load images (dataUrl → HTMLImageElement): picture elements + picture fills + background images
  // (recursing into group children and decoration nodes). Walk all pages so thumbnails get
  // background images/pictures of non-current pages (otherwise unvisited pages' thumbnails are blank).
  useEffect(() => {
    if (slides.length === 0) return
    const urls = new Set<string>()
    const addFillUrl = (fill: RenderFill | undefined) => {
      if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
    }
    const walk = (nodes: readonly RenderNode[]) => {
      for (const n of nodes) {
        if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
        if ((n.type === 'shape' || n.type === 'text') && n.fill) addFillUrl(n.fill)
        if (n.type === 'chart') addFillUrl((n as { bgFill?: RenderFill }).bgFill)
        if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
        if (n.type === 'table' && Array.isArray(n.cells)) {
          for (const c of n.cells) if (c.fill) addFillUrl(c.fill)
        }
      }
    }
    for (const s of slides) {
      addFillUrl(s.background)
      walk(s.nodes)
    }
    if (!imageLoaderRef.current) {
      imageLoaderRef.current = createImageLoader((entries) => {
        setImages((prev) => {
          const m = new Map(prev)
          for (const [k, v] of entries) m.set(k, v)
          return m
        })
      })
    }
    imageLoaderRef.current.load(urls)
  }, [slides])
  // Dispose on unmount and clear the ref so a remount (e.g. React Strict Mode's
  // dev double-mount) lazily recreates a fresh loader instead of reusing a disposed one.
  useEffect(
    () => () => {
      imageLoaderRef.current?.dispose()
      imageLoaderRef.current = null
    },
    [],
  )

  const editNode = useMemo(() => {
    if (!editing || !slide) return null
    // In-group-editing children: compose the group offset into an absolute box (the canvas only allows text editing when the group is unrotated/unflipped/unscaled)
    if (editing.groupId) {
      const g = slide.nodes.find((n) => n.sourceId === editing.groupId)
      if (g?.type !== 'group') return null
      const c = (g as GroupRenderNode).children.find((x) => x.sourceId === editing.sourceId)
      if (!c) return null
      return {
        ...c,
        box: { ...c.box, x: g.box.x + c.box.x, y: g.box.y + c.box.y },
      } as ShapeRenderNode
    }
    return slide.nodes.find((n) => n.sourceId === editing.sourceId) as ShapeRenderNode | undefined
  }, [editing, slide])

  const startEdit = useCallback(
    (sourceId: string, caret?: { x: number; y: number }) => {
      const isChild = enteredGroupNode?.children.some((c) => c.sourceId === sourceId)
      setEditing({ sourceId, caret, ...(isChild ? { groupId: enteredGroupNode!.sourceId } : {}) })
    },
    [enteredGroupNode],
  )

  // Audio/video playback overlay: triggered by double-clicking a media element, closed on page switch/Escape
  const [mediaPlay, setMediaPlay] = useState<{
    sourceId: string
    kind: 'video' | 'audio'
    dataUrl: string
  } | null>(null)
  useEffect(() => setMediaPlay(null), [current])
  const startMediaPlayback = useCallback(
    async (sourceId: string) => {
      const r = await window.slidesApi.getMediaData(current, sourceId)
      if (r) setMediaPlay({ sourceId, ...r })
    },
    [current],
  )
  const mediaPlayNode = useMemo(
    () =>
      mediaPlay ? (slide?.nodes.find((n) => n.sourceId === mediaPlay.sourceId) ?? null) : null,
    [mediaPlay, slide],
  )
  useEffect(() => {
    if (!mediaPlay) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setMediaPlay(null)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [mediaPlay])

  const startEditCell = useCallback((sourceId: string, row: number, col: number) => {
    setEditing(null)
    setEditingCell({ sourceId, row, col })
  }, [])

  /** Pseudo-node for the cell edit overlay: table box + cell local box → absolute page box */
  const cellEditNode = useMemo(() => {
    if (!editingCell || !slide) return null
    const tbl = slide.nodes.find((n) => n.sourceId === editingCell.sourceId)
    if (!tbl || tbl.type !== 'table') return null
    const cell = tbl.cells.find((c) => c.row === editingCell.row && c.col === editingCell.col)
    if (!cell) return null
    return {
      id: `cellEdit_${tbl.sourceId}_${cell.row}_${cell.col}`,
      type: 'shape',
      sourceId: tbl.sourceId,
      box: {
        x: tbl.box.x + cell.x,
        y: tbl.box.y + cell.y,
        w: cell.w,
        h: cell.h,
        rotationDeg: 0,
        flipH: false,
        flipV: false,
      },
      fill: { kind: 'none' },
      ...(cell.text ? { text: cell.text } : {}),
    } as unknown as ShapeRenderNode
  }, [editingCell, slide])

  // Table editing (table-actions.ts)
  const onTableColResize = useCallback(
    (sourceId: string, col: number, wPx: number) =>
      tableActions.onTableColResize(ctxRef.current, sourceId, col, wPx),
    [],
  )
  const onTableRowResize = useCallback(
    (sourceId: string, row: number, hPx: number) =>
      tableActions.onTableRowResize(ctxRef.current, sourceId, row, hPx),
    [],
  )
  const commitCellEdit = useCallback(
    (paragraphs: EditParagraph[]) => tableActions.commitCellEdit(ctxRef.current, paragraphs),
    [],
  )
  const navigateCell = useCallback(
    (paragraphs: EditParagraph[] | null, dir: 1 | -1) =>
      tableActions.navigateCell(ctxRef.current, paragraphs, dir),
    [],
  )

  const commitEdit = useCallback(
    async (paragraphs: EditParagraph[]) => {
      if (!editing) return
      const updated = await window.slidesApi.editText({
        slideIndex: current,
        sourceId: editing.sourceId,
        paragraphs,
        ...(editing.groupId ? { groupId: editing.groupId } : {}),
      })
      if (updated) {
        setSlides((s) => s.map((sl, i) => (i === current ? updated : sl)))
        setDirty(true)
      }
      setEditing(null)
      setSelectedIds([editing.sourceId]) // Back to shape-selected state after committing
    },
    [editing, current],
  )

  // ⌘+click on a linked run while editing: jump in the editor / open externally (same routing as the show)
  const followRunLink = useCallback(
    (target: LinkTargetOp) => {
      if (target.kind === 'slide') {
        if (target.slideIndex >= 0 && target.slideIndex < slides.length) {
          setCurrent(target.slideIndex)
        }
        return
      }
      window.open(target.url, '_blank', 'noreferrer')
    },
    [slides.length],
  )

  const onTransform = useCallback(
    async (
      sourceId: string,
      box: { x: number; y: number; w: number; h: number; rotationDeg: number },
      preview?: boolean,
      groupId?: string,
    ) => {
      const updated = await window.slidesApi.editTransform({
        slideIndex: current,
        sourceId,
        xPx: box.x,
        yPx: box.y,
        wPx: box.w,
        hPx: box.h,
        rotationDeg: box.rotationDeg,
        fitWidthPx: FIT_WIDTH,
        ...(preview ? { preview: true } : {}),
        ...(groupId ? { groupId } : {}),
      })
      if (updated) {
        setSlides((s) => s.map((sl, i) => (i === current ? updated : sl)))
        setDirty(true)
      }
    },
    [current],
  )

  // Connector endpoint drag: new endpoints + attach/detach for the dragged end
  const onEditConnectorEndpoints = useCallback(
    async (
      sourceId: string,
      ep: {
        x1: number
        y1: number
        x2: number
        y2: number
        start?: { targetId: string; idx: number } | null
        end?: { targetId: string; idx: number } | null
      },
    ) => {
      const updated = await window.slidesApi.editConnectorEndpoints({
        slideIndex: current,
        sourceId,
        x1Px: ep.x1,
        y1Px: ep.y1,
        x2Px: ep.x2,
        y2Px: ep.y2,
        fitWidthPx: FIT_WIDTH,
        ...(ep.start !== undefined ? { start: ep.start } : {}),
        ...(ep.end !== undefined ? { end: ep.end } : {}),
      })
      if (updated) {
        setSlides((s) => s.map((sl, i) => (i === current ? updated : sl)))
        setDirty(true)
      }
    },
    [current],
  )

  // Text / element styling (style-actions.ts)
  const onFormat = useCallback((cmd: FormatCmd) => styleActions.onFormat(cmd), [])
  const onFontFamily = useCallback(
    (family: string) => styleActions.onFontFamily(ctxRef.current, family),
    [],
  )
  const onFontSize = useCallback((pt: number) => styleActions.onFontSize(ctxRef.current, pt), [])

  // While editing, track font and size at the caret/selection (ribbon font group display)
  const [selFont, setSelFont] = useState<{ family: string; sizePt: number } | null>(null)
  // Paragraph alignment at the editing caret/selection (overlay DOM); null outside editing
  const [selAlign, setSelAlign] = useState<ParaAlign | null>(null)
  const inTextEdit = !!editing || !!editingCell
  useEffect(() => {
    if (!inTextEdit) {
      setSelFont(null)
      setSelAlign(null)
      return
    }
    const update = () => {
      const focus = window.getSelection()?.focusNode
      const el = focus instanceof HTMLElement ? focus : focus?.parentElement
      if (!el?.isContentEditable) return
      setSelAlign(liveAlign() ?? null)
      const cs = window.getComputedStyle(el)
      // Prefer the model font name baked into the run container (data-font) when the display font
      // is unchanged; the computed style may be a fallback/substitution product (e.g. Arial for DengXian)
      const container = el.closest('[data-display-font]') as HTMLElement | null
      const displayed = firstFontFamily(cs.fontFamily)
      const family =
        container?.dataset.font &&
        container.dataset.displayFont?.toLowerCase() === displayed.toLowerCase()
          ? container.dataset.font
          : displayed
      // The editor DOM is in viewport px (including scale/autofit fontScale); divide by norm to get model pt
      const root = el.closest('[contenteditable="true"]') as HTMLElement | null
      const norm = parseFloat(root?.dataset.norm ?? '') || 1
      const sizePt = Math.round(((parseFloat(cs.fontSize) * 72) / (96 * norm)) * 2) / 2
      setSelFont((v) => (v && v.family === family && v.sizePt === sizePt ? v : { family, sizePt }))
    }
    update()
    document.addEventListener('selectionchange', update)
    return () => document.removeEventListener('selectionchange', update)
  }, [inTextEdit])

  const onTextColor = useCallback((hex: string) => {
    document.execCommand('styleWithCSS', false, 'true')
    document.execCommand('foreColor', false, hex)
  }, [])

  const onAlign = useCallback((align: ParaAlign) => {
    styleActions.onAlign(ctxRef.current, align)
    // execCommand mutates only the overlay DOM (no state change, selectionchange isn't
    // guaranteed) — re-read so the ribbon highlight follows the click immediately
    if (ctxRef.current.editing || ctxRef.current.editingCell) setSelAlign(liveAlign() ?? null)
  }, [])
  const onTextToggle = useCallback(
    (kind: 'bold' | 'italic' | 'underline' | 'strike') =>
      styleActions.onTextToggle(ctxRef.current, kind),
    [],
  )
  const onStrike = useCallback(() => onTextToggle('strike'), [onTextToggle])
  const onElementTextColor = useCallback(
    (hex: string) => styleActions.onElementTextColor(ctxRef.current, hex),
    [],
  )
  // In-edit paragraph formats mutate the overlay DOM without a state change; the tick
  // re-runs curBulletChar so the gallery highlight follows the live marks
  const [paraFmtTick, setParaFmtTick] = useState(0)
  const onParagraphFormat = useCallback(
    (patch: Parameters<typeof styleActions.onParagraphFormat>[1]) => {
      styleActions.onParagraphFormat(ctxRef.current, patch)
      if (ctxRef.current.editing || ctxRef.current.editingCell) setParaFmtTick((n) => n + 1)
    },
    [],
  )
  const onFill = useCallback(
    (sourceId: string, fill: string | GradientFillSpec) =>
      styleActions.onFill(ctxRef.current, sourceId, fill),
    [],
  )

  const selectedNode = useMemo(
    () => (selectedIds.length === 1 ? (findNodeCtx(selectedIds[0]!)?.node ?? null) : null),
    [selectedIds, findNodeCtx],
  )

  const [selectedLink, setSelectedLink] = useState<LinkTargetOp | null>(null)
  useEffect(() => {
    if (selectedIds.length !== 1) {
      setSelectedLink(null)
      return
    }
    let alive = true
    void window.slidesApi.getLink(current, selectedIds[0]!).then((l) => {
      if (alive) setSelectedLink(l)
    })
    return () => {
      alive = false
    }
  }, [selectedIds, current, slides])

  const [selectedChartData, setSelectedChartData] = useState<Awaited<
    ReturnType<typeof window.slidesApi.getChartData>
  > | null>(null)
  useEffect(() => {
    if (selectedNode?.type !== 'chart') {
      setSelectedChartData(null)
      return
    }
    let alive = true
    void window.slidesApi.getChartData(current, selectedNode.sourceId).then((d) => {
      if (alive) setSelectedChartData(d)
    })
    return () => {
      alive = false
    }
  }, [selectedNode, current, slides])

  // Contextual tabs: expose the element type to the Ribbon for single selection
  const contextElementType = useMemo((): ContextElementType => {
    if (selectedNode) {
      return contextElementTypeForNode(selectedNode)
    }
    // Multi-select: all-shape selections get the shape-tools tab (styles/fill apply to
    // each); selections spanning pictures/groups keep picture-tools as 'mixed', where
    // outline applies to the whole selection and picture-only tools stay disabled
    if (selectedIds.length >= 2) {
      const nodes = selectedIds.map((id) => findNodeCtx(id)?.node)
      if (nodes.every((n) => n?.type === 'shape')) return 'shape'
      if (
        nodes.every((n) => n && (n.type === 'shape' || n.type === 'picture' || n.type === 'group'))
      )
        return 'mixed'
    }
    return null
  }, [selectedNode, selectedIds, findNodeCtx])

  /** Whether the selected picture supports background removal (audio/video poster frames / no data don't) */
  const contextPictureCanCutout = useMemo(() => {
    if (selectedNode?.type !== 'picture') return false
    const pic = selectedNode as PictureRenderNode
    return !pic.media && !!pic.dataUrl
  }, [selectedNode])

  const contextPictureStroke = useMemo(() => {
    // Multi-select shows the first shape/picture's current outline as the panel state
    // (for a group, the first strokeable child stands in)
    const strokeable = (n: RenderNode | undefined): PictureRenderNode | undefined => {
      if (!n) return undefined
      if (n.type === 'picture' || n.type === 'shape') return n as PictureRenderNode
      if (n.type === 'group')
        return (n as GroupRenderNode).children.find(
          (c) => c.type === 'picture' || c.type === 'shape',
        ) as PictureRenderNode | undefined
      return undefined
    }
    let s: PictureRenderNode['stroke']
    for (const id of selectedIds) {
      s = strokeable(findNodeCtx(id)?.node)?.stroke
      if (s) break
    }
    return s ? { color: s.color, widthPt: s.widthPt, dashPreset: s.dashPreset } : null
  }, [selectedIds, findNodeCtx])

  const contextChartStyle = useMemo(
    () =>
      selectedNode?.type === 'chart' ? ((selectedNode as ChartRenderNode).styleInfo ?? null) : null,
    [selectedNode],
  )

  const tableStyleFlags = useMemo(
    () =>
      selectedNode?.type === 'table'
        ? ((selectedNode as TableRenderNode).styleFlags ?? null)
        : null,
    [selectedNode],
  )

  const tableActiveCell = useMemo(
    () =>
      editingCell &&
      selectedNode?.type === 'table' &&
      selectedNode.sourceId === editingCell.sourceId
        ? { row: editingCell.row, col: editingCell.col }
        : null,
    [editingCell, selectedNode],
  )

  // Theme-derived chart colors (fetched when a chart is selected; re-selecting after applying a theme refreshes)
  const [chartColorSchemes, setChartColorSchemes] = useState<Array<{
    key: string
    label: string
    colors: string[]
  }> | null>(null)
  useEffect(() => {
    if (selectedNode?.type !== 'chart') return
    let stale = false
    // Optional call: degrade to the fixed palette when the preload build is too old, no blank screen
    void window.slidesApi.getChartColorSchemes?.().then((r) => {
      if (!stale && r) setChartColorSchemes(r)
    })
    return () => {
      stale = true
    }
  }, [selectedNode])

  const onEditTableStyle = useCallback(
    (op: Omit<EditTableStyleOp, 'slideIndex' | 'sourceId'>) =>
      styleActions.onEditTableStyle(ctxRef.current, op),
    [],
  )
  const onEditChart = useCallback(
    (op: Omit<EditChartOp, 'slideIndex' | 'sourceId'>) =>
      styleActions.onEditChart(ctxRef.current, op),
    [],
  )

  /** Open the chart data edit dialog */
  const [chartDataDialogOpen, setChartDataDialogOpen] = useState(false)
  const [chartDataDialogInit, setChartDataDialogInit] = useState<{
    kind: string
    title: string
    categories: string[]
    series: Array<{ name: string; values: number[] }>
  } | null>(null)

  const openChartDataDialog = useCallback(
    () => styleActions.openChartDataDialog(ctxRef.current),
    [],
  )
  // While editing, measure from the selection; with elements selected (incl. multi-select/groups/tables), aggregate all runs —
  // mixed fonts show empty, mixed sizes show the minimum plus "+"; elements without text (pictures/charts etc.) keep the last display;
  // final fallback is the theme body default font.
  // Whether the selection contains text-capable elements (text boxes/shapes/tables) — font group availability (pictures/charts etc. grayed)
  const hasTextSelection = useMemo(
    () =>
      selectedIds.some((id) => {
        const n = slide?.nodes.find((x) => x.sourceId === id)
        return !!n && (n.type === 'shape' || n.type === 'text' || n.type === 'table')
      }),
    [selectedIds, slide],
  )

  const lastFontRef = useRef<{ family: string; sizePt: number; sizeMixed?: boolean } | null>(null)
  const fontStatus = useMemo(() => {
    let st: { family: string; sizePt: number; sizeMixed?: boolean } | null = null
    if (inTextEdit) {
      st = selFont
    } else if (selectedIds.length) {
      const runs: Array<{ family: string; sizePt: number }> = []
      for (const n of slide?.nodes ?? []) {
        if (selectedIds.includes(n.sourceId)) collectFontRuns(n, slide?.scale ?? 1, runs)
      }
      if (runs.length) {
        const families = new Set(runs.map((r) => r.family))
        const sizes = runs.map((r) => r.sizePt)
        const minSize = Math.min(...sizes)
        st = {
          family: families.size === 1 ? runs[0]!.family : '',
          sizePt: minSize,
          sizeMixed: sizes.some((s) => s !== minSize),
        }
      } else {
        st = lastFontRef.current // Textless elements like pictures/charts: keep the last value
      }
    }
    if (!st && defaultFont) st = { family: defaultFont, sizePt: 18 }
    if (st) lastFontRef.current = st
    return st
  }, [inTextEdit, selFont, selectedIds, slide, defaultFont])

  // Current bullet char for the ribbon bullet gallery highlight: '' = no bullet
  // (None tile), a glyph = its preset tile, null = mixed/unknown (no highlight)
  const curBulletChar = useMemo(() => {
    void paraFmtTick // re-read the overlay DOM after an in-edit paragraph format
    if (inTextEdit) {
      // Uncommitted paragraph formats exist only in the overlay DOM, not the slide tree
      const live = liveBulletChar()
      if (live !== undefined) return live
    }
    if (editingCell) {
      // Cell edits scope to the one cell — the whole table would read as mixed
      const tbl = findNodeCtx(editingCell.sourceId)?.node
      if (tbl?.type !== 'table') return null
      const cell = tbl.cells.find((c) => c.row === editingCell.row && c.col === editingCell.col)
      if (!cell) return null
      const cellFound = new Set<string>()
      collectBodyBulletChars(cell.text, cellFound)
      return cellFound.size === 1 ? [...cellFound][0]! : null
    }
    const ids = selectedIds.length ? selectedIds : editing ? [editing.sourceId] : []
    if (!ids.length) return null
    const found = new Set<string>()
    // findNodeCtx also resolves children of the group being edited, not just top-level nodes
    for (const id of ids) {
      const node = findNodeCtx(id)?.node
      if (node) collectBulletChars(node, found)
    }
    if (found.size !== 1) return null
    return [...found][0]!
  }, [selectedIds, editing, editingCell, inTextEdit, findNodeCtx, paraFmtTick])

  // Current paragraph alignment for the ribbon highlight: unset text defaults to 'left',
  // so text always has exactly one alignment current; null = mixed/no text (no highlight)
  const curAlign = useMemo((): ParaAlign | null => {
    if (inTextEdit) return selAlign
    if (!selectedIds.length) return null
    const found = new Set<ParaAlign>()
    for (const id of selectedIds) {
      const node = findNodeCtx(id)?.node
      if (node) collectAligns(node, found)
    }
    return found.size === 1 ? [...found][0]! : null
  }, [inTextEdit, selAlign, selectedIds, findNodeCtx])

  // Refresh the action-module context every render so extracted actions never see stale state
  ctxRef.current = {
    slides,
    setSlides,
    current,
    setCurrent,
    slide,
    path,
    setPath,
    setDirty,
    setStatus,
    images,
    selectedIds,
    setSelectedIds,
    editing,
    setEditing,
    editingCell,
    setEditingCell,
    enteredGroupId,
    setEnteredGroupId,
    enteredGroupNode,
    selectedNode,
    hasClipboard,
    setHasClipboard,
    canPasteSlide,
    setCanPasteSlide,
    pasteFloater,
    setPasteFloater,
    brushFormat,
    setBrushFormat,
    brushMode,
    setBrushMode,
    animations,
    setAnimations,
    selAnim,
    setSelAnim,
    setHoverAnim,
    setTransition,
    animByParagraph,
    timingIdx,
    slideShow,
    setSlideShow,
    presenter,
    setPresenter,
    setCustomShows,
    setCustomShowDlgOpen,
    pendingRehearse,
    setPendingRehearse,
    sections,
    setSections,
    renamingSec,
    setRenamingSec,
    ctxMenu,
    setCtxMenu,
    cropTarget,
    setCropTarget,
    cutoutTarget,
    setCutoutTarget,
    linkDialog,
    setLinkDialog,
    setHfDialog,
    setEqDialogOpen,
    setChartDataDialogInit,
    setChartDataDialogOpen,
    setFindOpen,
    setPrintDlgOpen,
    setZoom,
    masterItems,
    recorderRef,
    setRecording,
    editingActiveRef,
    applySlide,
    flushNotes,
    findNodeCtx,
    groupIdOf,
    startEdit,
    undo,
    redo,
    onTransform,
    openBgFormat,
  }

  // Context menu items (context-menu-items.ts); the deps list covers all state the builder reads
  const ctxItems = useMemo(
    () => buildCtxItems(ctxRef.current),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [ctxMenu, slides, sections, hasClipboard, canPasteSlide, selectedIds, slide, lang],
  )

  // ---- Draw (freehand ink): one stroke = one transparent PNG picture element; undo/save use the existing pipeline ----

  const onInkTool = useCallback(
    (tool: InkTool) => {
      // Clicking the currently active tool again = cancel (back to select mode)
      const next = tool !== 'select' && inkTool === tool ? 'select' : tool
      setInkTool(next)
      if (next !== 'select') {
        setSelectedIds([])
        setEditing(null)
      }
    },
    [inkTool],
  )

  const commitInk = useCallback(
    (stroke: InkStroke) => arrangeActions.commitInk(ctxRef.current, stroke),
    [],
  )
  const eraseInk = useCallback(
    (sourceIds: string[]) => arrangeActions.eraseInk(ctxRef.current, sourceIds),
    [],
  )
  const inkCount = useMemo(() => (slide ? inkNodesOf(slide).length : 0), [slide])
  const clearInk = useCallback(() => arrangeActions.clearInk(ctxRef.current), [])

  const _fileName = slide ? path?.split('/').pop() || t('appUntitledPresentation') : undefined

  return (
    <div className="app">
      <ToastHost />
      <Ribbon
        hasDoc={!!slide}
        deckEmpty={deckEmpty}
        canUndo={histState.canUndo}
        canRedo={histState.canRedo}
        dirty={dirty}
        editing={!!editing || !!editingCell}
        autoSave={autoSave}
        onAutoSaveChange={setAutoSave}
        onOpen={() => void openDialog()}
        onSave={() => void save()}
        onUndo={() => void undo()}
        onRedo={() => void redo()}
        onSaveAs={() => void saveAs()}
        onExportPdf={() => void exportPdf()}
        onPrint={() => setPrintDlgOpen(true)}
        onExportImages={() => void exportImages()}
        onFormat={onFormat}
        zoom={zoom}
        onZoom={previewZoom}
        showThumbs={showThumbs}
        onToggleThumbs={() => setShowThumbs((v) => !v)}
        aiOpen={showAi}
        onToggleAi={toggleAi}
        onAiPreset={(text, opts) => pushAiPreset(text, true, undefined, undefined, opts?.slideShot)}
        onInsert={(kind) => void insertElement(kind)}
        onPickShape={pickShape}
        onInsertImage={() => void insertImage()}
        onFormatBackground={openBgFormat}
        onApplyTheme={(preset) => void applyThemePreset(preset)}
        onAddSlide={() => void addSlide()}
        onAddSection={() => void addSectionAt(current)}
        onAddSlideWithLayout={(lp) => void addSlideWithLayout(lp)}
        layouts={layoutsResult?.layouts ?? null}
        layoutSize={layoutsResult?.size ?? null}
        formatOpen={showFormat}
        onToggleFormat={() =>
          setShowFormat((v) => {
            if (!v) {
              setShowAnimPane(false)
              setShowBgFormat(false)
            }
            return !v
          })
        }
        hasSelection={selectedIds.length > 0}
        hasTextSelection={hasTextSelection}
        canPaste={hasClipboard}
        onCopy={() => void copySelected()}
        onCut={() => void cutSelected()}
        onPaste={() => void pasteClipboard()}
        hasBrushFormat={brushFormat !== null}
        brushMode={brushMode}
        onFormatBrushClick={onFormatBrushClick}
        onFormatBrushDoubleClick={onFormatBrushDoubleClick}
        onTextColor={onTextColor}
        onAlign={onAlign}
        onStrike={onStrike}
        onTextToggle={onTextToggle}
        onElementTextColor={onElementTextColor}
        onFindReplace={() => setFindOpen(true)}
        animByParagraph={animByParagraph}
        onToggleAnimByParagraph={() => setAnimByParagraph((v) => !v)}
        onSetLayout={(layoutPath) =>
          void window.slidesApi
            .setSlideLayout({ slideIndex: current, layoutPath })
            .then((r) => r && applySlide(current, r))
        }
        onResetLayout={() =>
          void window.slidesApi
            .setSlideLayout({ slideIndex: current })
            .then((r) => r && applySlide(current, r))
        }
        onSlideSize={(cx, cy) =>
          void window.slidesApi.setSlideSize({ cx, cy }).then((all) => {
            if (all) {
              setSlides(all)
              // Every slide re-materializes with fresh element ids
              setSelectedIds([])
              setEditing(null)
              setDirty(true)
            }
          })
        }
        slideSizeKey={
          slide
            ? Math.abs(slide.widthPx / slide.heightPx - 16 / 9) < 0.02
              ? '16:9'
              : Math.abs(slide.widthPx / slide.heightPx - 4 / 3) < 0.02
                ? '4:3'
                : null
            : null
        }
        onParagraphFormat={onParagraphFormat}
        curBulletChar={curBulletChar}
        curAlign={curAlign}
        curFontFamily={fontStatus?.family ?? null}
        curFontSizePt={fontStatus?.sizePt ?? null}
        curFontSizeMixed={fontStatus?.sizeMixed ?? false}
        onFontFamily={onFontFamily}
        onFontSize={onFontSize}
        onInsertTable={(rows, cols) => void insertTable(rows, cols)}
        transition={transition}
        onTransition={(kind, all) => {
          void applyTransition(kind, all)
          if (!all) previewTransitionOnCanvas(kind)
        }}
        selectedAnimEffect={selectedAnimEffect}
        timingAnim={timingIdx >= 0 ? animations[timingIdx]! : null}
        onApplyAnimation={applyAnimation}
        onAnimHoverPreview={hoverPreviewAnimation}
        onAnimHoverEnd={() => setHoverAnim(null)}
        onAddAnimation={addAnimation}
        onApplyMotionPath={applyMotionPath}
        onAnimTiming={patchAnimTiming}
        animPaneOpen={showAnimPane}
        onToggleAnimPane={toggleAnimPane}
        animCount={animations.length}
        onAnimPreview={() => setAnimPreview((n) => n + 1)}
        onSlideShow={startSlideShow}
        onPresenterView={startPresenterView}
        onCustomShow={() => setCustomShowDlgOpen(true)}
        onRehearse={startRehearseShow}
        currentHidden={!!slide?.hidden}
        onToggleHidden={() => void toggleHidden(current)}
        inkTool={inkTool}
        onInkTool={onInkTool}
        inkPen={inkPen}
        onInkPen={setInkPen}
        inkHighlighter={inkHighlighter}
        onInkHighlighter={setInkHighlighter}
        inkCount={inkCount}
        onInkClearAll={() => void clearInk()}
        viewMode={viewMode}
        onViewMode={onViewMode}
        onSlideMaster={() => void enterMasterView()}
        onZoomFit={zoomToFit}
        showRuler={showRuler}
        onToggleRuler={() => setShowRuler((v) => !v)}
        showGrid={showGrid}
        onToggleGrid={() => setShowGrid((v) => !v)}
        showGuides={showGuides}
        onToggleGuides={() => setShowGuides((v) => !v)}
        showNotes={showNotes}
        onToggleNotes={() => setShowNotes((v) => !v)}
        commentsOpen={showComments}
        onToggleComments={() => (showComments ? setShowComments(false) : openComments(false))}
        onNewComment={() => openComments(true)}
        commentCount={comments.length}
        onInsertIcon={(def, color) => void insertIcon(def, color)}
        onInsertChart={(kind) => void insertChart(kind)}
        onInsertSmartArt={(def) => void insertSmartArt(def)}
        onInsertWordArt={(preset) => void insertWordArt(preset)}
        onInsertField={(type) => void insertField(type)}
        onOpenLink={() => void openLinkDialog()}
        onInsertZoom={(index) => void insertZoom(index)}
        slideCount={slides.length}
        currentSlide={current}
        onOpenHeaderFooter={() => void openHeaderFooter()}
        onOpenEquation={() => setEqDialogOpen(true)}
        onInsertMedia={(kind) => void insertMediaFile(kind)}
        onInsertModel3d={() => void insertModel3dFile()}
        recording={recording}
        onToggleScreenRecord={() => void toggleScreenRecord()}
        contextElementType={contextElementType}
        contextElementId={selectedNode?.sourceId}
        contextSlideIndex={current}
        contextChartStyle={contextChartStyle}
        chartColorSchemes={chartColorSchemes}
        contextPictureCanCutout={contextPictureCanCutout}
        contextPictureStroke={contextPictureStroke}
        onPictureStroke={(stroke) => {
          // applies to every selected shape/picture; a selected group pierces one level
          // down so its shape/picture members get the outline (PowerPoint semantics)
          for (const id of selectedIds) {
            const n = findNodeCtx(id)?.node
            if (!n) continue
            if (n.type === 'picture' || n.type === 'shape') {
              void onStroke(id, stroke)
            } else if (n.type === 'group') {
              for (const c of (n as GroupRenderNode).children) {
                if (c.type === 'picture' || c.type === 'shape')
                  void window.slidesApi
                    .editStroke({
                      slideIndex: current,
                      sourceId: c.sourceId,
                      stroke,
                      groupId: n.sourceId,
                    })
                    .then((r) => r && applySlide(current, r))
              }
            }
          }
        }}
        onShapeStyle={(s) => {
          // fill + outline together, applied to every selected shape (sequentially: both
          // edits rewrite the same slide XML in the main process); a selected group
          // pierces one level down to its shape members (PowerPoint semantics)
          void (async () => {
            const applyTo = async (shape: ShapeRenderNode, groupId?: string) => {
              const st = shape.stroke
              const stroke = { color: s.stroke, widthPt: st?.widthPt ?? 1, dash: s.dash ?? 'solid' }
              if (groupId) {
                const r1 = await window.slidesApi.editFill({
                  slideIndex: current,
                  sourceId: shape.sourceId,
                  fill: s.fill,
                  groupId,
                })
                if (r1) applySlide(current, r1)
                const r2 = await window.slidesApi.editStroke({
                  slideIndex: current,
                  sourceId: shape.sourceId,
                  stroke,
                  groupId,
                })
                if (r2) applySlide(current, r2)
              } else {
                await onFill(shape.sourceId, s.fill)
                await onStroke(shape.sourceId, stroke)
              }
            }
            for (const id of selectedIds) {
              const n = findNodeCtx(id)?.node
              if (n?.type === 'shape') {
                await applyTo(n as ShapeRenderNode)
              } else if (n?.type === 'group') {
                for (const c of (n as GroupRenderNode).children) {
                  if (c.type === 'shape') await applyTo(c as ShapeRenderNode, n.sourceId)
                }
              }
            }
          })()
        }}
        onShapeFill={(fill) => {
          void (async () => {
            for (const id of selectedIds) {
              const n = findNodeCtx(id)?.node
              if (n?.type === 'shape') {
                await onFill(id, fill)
              } else if (n?.type === 'group') {
                for (const c of (n as GroupRenderNode).children) {
                  if (c.type !== 'shape') continue
                  const r = await window.slidesApi.editFill({
                    slideIndex: current,
                    sourceId: c.sourceId,
                    fill,
                    groupId: n.sourceId,
                  })
                  if (r) applySlide(current, r)
                }
              }
            }
          })()
        }}
        onPictureCrop={startCrop}
        cropActive={cropTarget != null}
        onPictureCutout={startCutout}
        onPictureOpacity={(opacity) => {
          if (!selectedNode || selectedNode.type !== 'picture') return
          void window.slidesApi
            .editPictureOpacity({ slideIndex: current, sourceId: selectedNode.sourceId, opacity })
            .then((r) => r && applySlide(current, r))
        }}
        onEditTableStyle={(op) => void onEditTableStyle(op)}
        tableStyleFlags={tableStyleFlags}
        tableActiveCell={tableActiveCell}
        onEditChart={(op) => void onEditChart(op)}
        onOpenChartDataDialog={() => void openChartDataDialog()}
        onArrange={(op) => void alignSelected(op)}
        onFlip={(axis) => void flipSelected(axis)}
        canDistribute={selectedIds.length >= 3}
      />

      <div className="app-main">
        {slide && viewMode !== 'reading' && viewMode !== 'sorter' && (
          <div className={`ai-dock${showAi && aiSettings ? '' : ' collapsed'}`}>
            {/* always mounted once settings load: collapse must not drop state or in-flight runs */}
            {aiSettings ? (
              <AiPanel
                key={aiPanelKey}
                slides={slides}
                current={current}
                selectedIds={selectedIds}
                deckEmpty={deckEmpty}
                images={images}
                applySlide={applySlide}
                applyDeck={applyDeck}
                fitWidthPx={FIT_WIDTH}
                settings={aiSettings}
                preset={aiPreset}
                open={showAi}
                onExpand={toggleAi}
                onCollapse={toggleAi}
                onUndo={() => void undo()}
                onPathChange={(p) => {
                  setPath(p)
                  setDirty(false)
                }}
                currentFilePath={path}
              />
            ) : (
              <button
                className="ai-rail"
                onClick={toggleAi}
                data-tip={t('appAiRailExpand')}
                aria-label={t('appAiRailExpand')}
              >
                <GensparkMark size={22} />
              </button>
            )}
          </div>
        )}
        <div className="app-content">
          <div className="workspace">
            {!slide ? (
              <div className="start-screen start-booting">{t('appStartOpening')}</div>
            ) : viewMode === 'reading' ? (
              (() => {
                // Reading view: page fills the available area, click/arrow keys to turn pages, Esc to exit
                const availH = Math.max(240, winSize.h - 40 - 112 - 24 - 76)
                const availW = Math.max(320, winSize.w - 120)
                const readW = Math.round(
                  Math.min(availW, availH * (slide.widthPx / slide.heightPx)),
                )
                return (
                  <div className="reading-view">
                    <div
                      className="reading-slide"
                      data-tip={t('appReadingSlideTitle')}
                      onClick={() => setCurrent((c) => Math.min(c + 1, slides.length - 1))}
                    >
                      <SlideThumb slide={slide} images={images} width={readW} />
                    </div>
                    <div className="reading-bar">
                      <button
                        disabled={current === 0}
                        onClick={() => setCurrent((c) => Math.max(c - 1, 0))}
                      >
                        {t('appReadingPrev')}
                      </button>
                      <span className="reading-page">
                        {current + 1} / {slides.length}
                      </span>
                      <button
                        disabled={current === slides.length - 1}
                        onClick={() => setCurrent((c) => Math.min(c + 1, slides.length - 1))}
                      >
                        {t('appReadingNext')}
                      </button>
                      <button className="reading-exit" onClick={() => onViewMode('normal')}>
                        {t('appReadingExit')}
                      </button>
                    </div>
                  </div>
                )
              })()
            ) : viewMode === 'sorter' ? (
              <div className="sorter-view">
                {slides.map((s, i) => (
                  <div
                    key={i}
                    className={`sorter-item ${i === current ? 'active' : ''} ${s.hidden ? 'thumb-hidden' : ''}${thumbDragCls(i)}`}
                    data-tip={s.hidden ? t('appSorterHiddenTitle') : t('appSorterItemTitle')}
                    {...thumbDragProps(i, true)}
                    onClick={() => {
                      setCurrent(i)
                      setSelectedIds([])
                      setEditing(null)
                    }}
                    onDoubleClick={() => {
                      setCurrent(i)
                      onViewMode('normal')
                    }}
                    onContextMenu={(e) => {
                      e.preventDefault()
                      setCurrent(i)
                      void window.slidesApi.hasSlideClipboard().then(setCanPasteSlide)
                      setCtxMenu({ kind: 'thumb', x: e.clientX, y: e.clientY, index: i })
                    }}
                  >
                    <SlideThumb slide={s} images={images} width={208} />
                    <span className="sorter-num">{i + 1}</span>
                    {pasteFloater?.index === i && (
                      <PasteOptionsFloater
                        mode={pasteFloater.mode}
                        onSelect={repasteSlideAs}
                        onDismiss={() => setPasteFloater(null)}
                      />
                    )}
                  </div>
                ))}
              </div>
            ) : (
              <>
                {viewMode === 'outline' ? (
                  <div className="outline-pane">
                    {slides.map((s, i) => {
                      const o = outlineOf(s)
                      return (
                        <div
                          key={i}
                          className={`outline-item ${i === current ? 'active' : ''}`}
                          onClick={() => {
                            setCurrent(i)
                            setSelectedIds([])
                            setEditing(null)
                          }}
                        >
                          <span className="outline-num">{i + 1}</span>
                          <div className="outline-body">
                            <div className="outline-title">{o.title || t('appOutlineNoText')}</div>
                            {o.lines.slice(0, 5).map((t, j) => (
                              <div key={j} className="outline-line">
                                {t}
                              </div>
                            ))}
                            {o.lines.length > 5 && (
                              <div className="outline-line outline-more">…</div>
                            )}
                          </div>
                        </div>
                      )
                    })}
                  </div>
                ) : (
                  showThumbs && (
                    <>
                      <div className="slide-list" ref={thumbsListRef} style={{ width: thumbsW }}>
                        {(() => {
                          // width = sidebar minus horizontal padding (20) and .thumb border (4)
                          const thumbW = Math.max(60, thumbsW - 24)
                          const thumbItem = (s: RenderSlide, i: number) => (
                            <div
                              key={i}
                              className={`thumb ${i === current ? 'active' : ''} ${s.hidden ? 'thumb-hidden' : ''}${thumbDragCls(i)}`}
                              data-tip={s.hidden ? t('appThumbHiddenTitle') : undefined}
                              {...thumbDragProps(i)}
                              onClick={() => {
                                setCurrent(i)
                                setSelectedIds([])
                                setEditing(null)
                              }}
                              onContextMenu={(e) => {
                                e.preventDefault()
                                setCurrent(i)
                                setSelectedIds([])
                                setEditing(null)
                                void window.slidesApi.hasSlideClipboard().then(setCanPasteSlide)
                                setCtxMenu({ kind: 'thumb', x: e.clientX, y: e.clientY, index: i })
                              }}
                            >
                              <SlideThumb slide={s} images={images} width={thumbW} />
                              <span className="thumb-num">{i + 1}</span>
                              {pasteFloater?.index === i && (
                                <PasteOptionsFloater
                                  mode={pasteFloater.mode}
                                  onSelect={repasteSlideAs}
                                  onDismiss={() => setPasteFloater(null)}
                                />
                              )}
                            </div>
                          )
                          if (!sectionGroups) return slides.map((s, i) => thumbItem(s, i))
                          return sectionGroups.map((g, gi) => {
                            const collapsed = g.id != null && collapsedSecs.has(g.id)
                            return (
                              <div key={g.id ?? `lead-${gi}`} className="section-group">
                                <div
                                  className={`section-header${g.id == null ? ' section-header-none' : ''}`}
                                  onClick={g.id != null ? () => toggleSection(g.id!) : undefined}
                                  onContextMenu={
                                    g.id != null
                                      ? (e) => {
                                          e.preventDefault()
                                          setCtxMenu({
                                            kind: 'section',
                                            x: e.clientX,
                                            y: e.clientY,
                                            sectionId: g.id!,
                                          })
                                        }
                                      : undefined
                                  }
                                  title={g.id != null ? t('appSectionHeaderTitle') : undefined}
                                >
                                  {g.id != null && (
                                    <span className={`section-arrow${collapsed ? '' : ' open'}`}>
                                      ▸
                                    </span>
                                  )}
                                  {renamingSec && g.id != null && renamingSec.id === g.id ? (
                                    <input
                                      className="section-rename-input"
                                      autoFocus
                                      value={renamingSec.value}
                                      onClick={(e) => e.stopPropagation()}
                                      onChange={(e) =>
                                        setRenamingSec({ id: g.id!, value: e.target.value })
                                      }
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') commitRenameSection()
                                        else if (e.key === 'Escape') setRenamingSec(null)
                                      }}
                                      onBlur={commitRenameSection}
                                    />
                                  ) : (
                                    <span
                                      className="section-name"
                                      onDoubleClick={
                                        g.id != null
                                          ? () => setRenamingSec({ id: g.id!, value: g.name })
                                          : undefined
                                      }
                                    >
                                      {g.name}
                                    </span>
                                  )}
                                  <span className="section-count">{g.end - g.start}</span>
                                </div>
                                {!collapsed &&
                                  slides
                                    .slice(g.start, g.end)
                                    .map((s, k) => thumbItem(s, g.start + k))}
                              </div>
                            )
                          })
                        })()}
                      </div>
                      <div className="thumb-resizer" onPointerDown={startThumbsResize} />
                    </>
                  )
                )}
                <div className="stage-col">
                  {showRuler && (
                    <div className="ruler-row">
                      <div className="ruler-corner" />
                      <StageRuler
                        unit={rulerUnit}
                        zoom={zoom}
                        slideLen={slide.widthPx}
                        origin={rulerOrigin?.x ?? 0}
                        wrapRef={stageWrapRef}
                        onGuidePreview={(client) => {
                          const pos = Math.min(1, Math.max(0, guideFracAt('h', client)))
                          setGuidePreview({ axis: 'h', pos })
                          setGuideBubble({
                            x: client.x,
                            y: client.y,
                            text: formatRulerValue(pos, slide.heightPx, rulerUnit),
                          })
                        }}
                        onGuideCommit={(client) => {
                          setGuidePreview(null)
                          setGuideBubble(null)
                          // released back over the chrome (outside the slide) = cancel
                          const raw = guideFracAt('h', client)
                          if (raw < -0.02 || raw > 1.02) return
                          setGuides((g) => [
                            ...g,
                            { axis: 'h', pos: Math.min(1, Math.max(0, raw)) },
                          ])
                          setShowGuides(true)
                        }}
                      />
                    </div>
                  )}
                  <div className="ruler-track">
                    {showRuler && (
                      <StageRuler
                        vertical
                        unit={rulerUnit}
                        zoom={zoom}
                        slideLen={slide.heightPx}
                        origin={rulerOrigin?.y ?? 0}
                        wrapRef={stageWrapRef}
                        onGuidePreview={(client) => {
                          const pos = Math.min(1, Math.max(0, guideFracAt('v', client)))
                          setGuidePreview({ axis: 'v', pos })
                          setGuideBubble({
                            x: client.x,
                            y: client.y,
                            text: formatRulerValue(pos, slide.widthPx, rulerUnit),
                          })
                        }}
                        onGuideCommit={(client) => {
                          setGuidePreview(null)
                          setGuideBubble(null)
                          // released back over the chrome (outside the slide) = cancel
                          const raw = guideFracAt('v', client)
                          if (raw < -0.02 || raw > 1.02) return
                          setGuides((g) => [
                            ...g,
                            { axis: 'v', pos: Math.min(1, Math.max(0, raw)) },
                          ])
                          setShowGuides(true)
                        }}
                      />
                    )}
                    <div
                      className={`stage-wrap${stageFitsViewport ? ' stage-fits-viewport' : ''}${brushMode ? ' format-brush-mode' : ''}`}
                      ref={stageWrapRef}
                    >
                      {/* transform: scale() doesn't grow layout, so the scroll range ignores the
                    zoomed size and the left/top overflow becomes unreachable; the zoom-box
                    is sized to the scaled dimensions to give the scroller the real extent */}
                      <div
                        ref={zoomBoxRef}
                        className="stage-zoom-box"
                        style={
                          scaleBox
                            ? { width: scaleBox.w * zoom, height: scaleBox.h * zoom }
                            : undefined
                        }
                      >
                        <div className="stage-ai-bar">
                          <button
                            className={`stage-ai-btn${showAi ? ' active' : ''}`}
                            data-tip={t('aiOpenAssistant')}
                            onClick={toggleAi}
                          >
                            <GensparkMark size={14} />
                            <span>Genspark AI</span>
                          </button>
                          {/* Same one-click presets as the Home tab; hidden instead of
                        disabled while the deck has no real content */}
                          {!deckEmpty && (
                            <>
                              <span className="stage-ai-divider" aria-hidden="true" />
                              <button
                                className="stage-ai-btn"
                                data-tip={t('aiBeautifyBtn')}
                                onClick={() =>
                                  pushAiPreset(
                                    t('aiBeautifyPrompt'),
                                    true,
                                    undefined,
                                    undefined,
                                    true,
                                  )
                                }
                              >
                                <IconAiBeautify size={14} />
                                <span>{t('aiBeautifyBtn')}</span>
                              </button>
                              <button
                                className="stage-ai-btn"
                                data-tip={t('aiFactCheckBtn')}
                                onClick={() => pushAiPreset(t('aiFactCheckPrompt'))}
                              >
                                <IconAiFactCheck size={14} />
                                <span>{t('aiFactCheckBtn')}</span>
                              </button>
                              <button
                                className="stage-ai-btn"
                                data-tip={t('aiImageBtn')}
                                onClick={() => pushAiPreset(t('aiImagePrompt'))}
                              >
                                <IconAiImage size={14} />
                                <span>{t('aiImageBtn')}</span>
                              </button>
                            </>
                          )}
                        </div>
                        <div
                          ref={stageScaleRef}
                          className="stage-scale"
                          style={{ transform: `scale(${zoom})`, transformOrigin: 'top left' }}
                        >
                          <div
                            ref={stageRelRef}
                            className={`stage-rel${transPreviewKind ? ` tp-${transPreviewKind}` : ''}`}
                            onAnimationEnd={(e) => {
                              if (e.target === e.currentTarget) setTransPreviewKind(null)
                            }}
                            style={{
                              position: 'relative',
                              width: slide.widthPx,
                              height: slide.heightPx,
                            }}
                            onDragOver={(e) => {
                              if (e.dataTransfer.types.includes('Files')) e.preventDefault()
                            }}
                            onDrop={(e) => {
                              const files = Array.from(e.dataTransfer.files).filter((f) =>
                                f.type.startsWith('image/'),
                              )
                              if (!files.length) return
                              e.preventDefault()
                              const rect = e.currentTarget.getBoundingClientRect()
                              const at = {
                                x: ((e.clientX - rect.left) / rect.width) * slide.widthPx,
                                y: ((e.clientY - rect.top) / rect.height) * slide.heightPx,
                              }
                              for (const f of files) {
                                void f.arrayBuffer().then((buf) => {
                                  // Chunked base64 conversion: spreading a large array would blow the call stack
                                  const bytes = new Uint8Array(buf)
                                  let bin = ''
                                  for (let i = 0; i < bytes.length; i += 0x8000) {
                                    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
                                  }
                                  const ext = (f.name.split('.').pop() ?? 'png').toLowerCase()
                                  void insertExternalImage(
                                    btoa(bin),
                                    ext === 'jpeg' ? 'jpg' : ext,
                                    at,
                                  )
                                })
                              }
                            }}
                          >
                            <SlideCanvas
                              slide={slide}
                              selectedIds={selectedIds}
                              onSelect={handleCanvasSelect}
                              onEditText={startEdit}
                              onTransform={onTransform}
                              onEditTableCell={startEditCell}
                              onTableColResize={onTableColResize}
                              onTableRowResize={(id, row, hPx) =>
                                void onTableRowResize(id, row, hPx)
                              }
                              onPlayMedia={(id) => void startMediaPlayback(id)}
                              onContextMenu={onCanvasContextMenu}
                              onMarqueeSelect={setSelectedIds}
                              onDuplicateTo={(id, dx, dy) => void duplicateSelected([id], dx, dy)}
                              images={images}
                              zoom={zoom}
                              enteredGroupId={enteredGroupId}
                              onEnterGroup={handleEnterGroup}
                              onEditConnectorEndpoints={onEditConnectorEndpoints}
                              drawMode={drawKind ? { kind: drawKind } : null}
                              onDrawCommit={commitDraw}
                              onDrawCancel={() => {
                                drawKindRef.current = null
                                setDrawKind(null)
                              }}
                              editingText={
                                editing
                                  ? { sourceId: editing.sourceId }
                                  : editingCell
                                    ? {
                                        sourceId: editingCell.sourceId,
                                        cell: { row: editingCell.row, col: editingCell.col },
                                      }
                                    : null
                              }
                            />
                            {showGrid && <div className="grid-overlay" />}
                            {showGuides &&
                              guides.map((g, gi) => (
                                <div
                                  key={`${g.axis}${gi}`}
                                  className={`guide-line guide-${g.axis} guide-draggable`}
                                  style={
                                    g.axis === 'v'
                                      ? { left: `${g.pos * 100}%` }
                                      : { top: `${g.pos * 100}%` }
                                  }
                                  data-tip={t('appGuideDragTip')}
                                  onPointerDown={(e) => {
                                    e.preventDefault()
                                    e.currentTarget.setPointerCapture(e.pointerId)
                                    const host = e.currentTarget.parentElement!
                                    const rect = host.getBoundingClientRect()
                                    const move = (ev: PointerEvent) => {
                                      const p = Math.min(
                                        1,
                                        Math.max(
                                          0,
                                          g.axis === 'v'
                                            ? (ev.clientX - rect.left) / rect.width
                                            : (ev.clientY - rect.top) / rect.height,
                                        ),
                                      )
                                      setGuides((list) =>
                                        list.map((x, i) => (i === gi ? { ...x, pos: p } : x)),
                                      )
                                      setGuideBubble({
                                        x: ev.clientX,
                                        y: ev.clientY,
                                        text: formatRulerValue(
                                          p,
                                          g.axis === 'v' ? slide.widthPx : slide.heightPx,
                                          rulerUnit,
                                        ),
                                      })
                                    }
                                    const up = (ev: PointerEvent) => {
                                      window.removeEventListener('pointermove', move)
                                      window.removeEventListener('pointerup', up)
                                      setGuideBubble(null)
                                      // Dragging off the canvas = delete the guide
                                      const outside =
                                        ev.clientX < rect.left - 24 ||
                                        ev.clientX > rect.right + 24 ||
                                        ev.clientY < rect.top - 24 ||
                                        ev.clientY > rect.bottom + 24
                                      if (outside)
                                        setGuides((list) => list.filter((_, i) => i !== gi))
                                    }
                                    window.addEventListener('pointermove', move)
                                    window.addEventListener('pointerup', up)
                                  }}
                                  onDoubleClick={() =>
                                    setGuides((list) => list.filter((_, i) => i !== gi))
                                  }
                                />
                              ))}
                            {guidePreview && (
                              <div
                                className={`guide-line guide-${guidePreview.axis}`}
                                style={
                                  guidePreview.axis === 'v'
                                    ? { left: `${guidePreview.pos * 100}%` }
                                    : { top: `${guidePreview.pos * 100}%` }
                                }
                              />
                            )}
                            {editing && editNode && (
                              <TextEditOverlay
                                node={editNode}
                                scale={slide.scale}
                                caretPoint={editing.caret}
                                replaceWith={editing.replaceWith}
                                onCommit={commitEdit}
                                onCancel={() => setEditing(null)}
                                onFollowLink={followRunLink}
                                frameColor={selectionChromeColor(slide, images)}
                                zoom={zoom}
                              />
                            )}
                            {editingCell && cellEditNode && (
                              <TextEditOverlay
                                node={cellEditNode}
                                scale={slide.scale}
                                onCommit={commitCellEdit}
                                onCancel={() => setEditingCell(null)}
                                onTabNav={(paragraphs, dir) => void navigateCell(paragraphs, dir)}
                                onFollowLink={followRunLink}
                                frameColor={selectionChromeColor(slide, images)}
                                zoom={zoom}
                              />
                            )}
                            {mediaPlay && mediaPlayNode && (
                              <div
                                style={{
                                  position: 'absolute',
                                  left: mediaPlayNode.box.x,
                                  top: mediaPlayNode.box.y,
                                  width: mediaPlayNode.box.w,
                                  height:
                                    mediaPlay.kind === 'video'
                                      ? mediaPlayNode.box.h
                                      : Math.max(mediaPlayNode.box.h, 40),
                                  zIndex: 20,
                                  background: mediaPlay.kind === 'video' ? '#000' : 'transparent',
                                  borderRadius: 4,
                                  overflow: 'hidden',
                                  boxShadow: '0 2px 12px rgba(0,0,0,.35)',
                                }}
                              >
                                {mediaPlay.kind === 'video' ? (
                                  <video
                                    src={mediaPlay.dataUrl}
                                    controls
                                    autoPlay
                                    style={{ width: '100%', height: '100%', objectFit: 'contain' }}
                                  />
                                ) : (
                                  <audio
                                    src={mediaPlay.dataUrl}
                                    controls
                                    autoPlay
                                    style={{ width: '100%' }}
                                  />
                                )}
                                <button
                                  onClick={() => setMediaPlay(null)}
                                  data-tip={t('appMediaCloseTitle')}
                                  aria-label={t('appMediaCloseTitle')}
                                  style={{
                                    position: 'absolute',
                                    top: 4,
                                    right: 4,
                                    width: 22,
                                    height: 22,
                                    lineHeight: '20px',
                                    padding: 0,
                                    border: 'none',
                                    borderRadius: '50%',
                                    background: 'rgba(0,0,0,.55)',
                                    color: '#fff',
                                    cursor: 'pointer',
                                  }}
                                >
                                  ×
                                </button>
                              </div>
                            )}
                            {animPreview > 0 && (
                              <AnimPreviewOverlay
                                key={animPreview}
                                slide={slide}
                                images={images}
                                items={animations}
                                onDone={() => setAnimPreview(0)}
                              />
                            )}
                            {hoverAnim && animPreview === 0 && (
                              <AnimPreviewOverlay
                                key={`hover-${hoverAnim.nonce}`}
                                slide={slide}
                                images={images}
                                items={hoverAnim.items}
                                onDone={() => setHoverAnim(null)}
                              />
                            )}
                            {cropTarget && (
                              <CropOverlay
                                box={cropTarget.box}
                                fullBox={cropTarget.fullBox}
                                imgSrc={cropTarget.dataUrl}
                                onConfirm={(rect) => void commitCrop(rect)}
                                onCancel={cancelCrop}
                              />
                            )}
                            {inkTool !== 'select' && !editing && (
                              <InkOverlay
                                slide={slide}
                                tool={inkTool}
                                color={
                                  inkTool === 'highlighter' ? inkHighlighter.color : inkPen.color
                                }
                                width={
                                  inkTool === 'highlighter' ? inkHighlighter.width : inkPen.width
                                }
                                onCommit={(stroke) => void commitInk(stroke)}
                                onErase={(ids) => void eraseInk(ids)}
                              />
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                  {guideBubble && (
                    <div
                      className="ruler-bubble"
                      style={{ left: guideBubble.x + 14, top: guideBubble.y + 16 }}
                    >
                      {guideBubble.text}
                    </div>
                  )}
                  {/* Notes pane, PowerPoint-style: shown by default as a thin
                      one-line bar you type into directly; drag the top edge to
                      any height (dragging to the very bottom hides it); the
                      ribbon/status 备注 buttons hide/show it entirely. When
                      hidden, an invisible strip along the bottom edge stays
                      grabbable to pull the pane back out, like PPT's splitter. */}
                  {!showNotes && (
                    <div className="notes-pull-zone" onMouseDown={(e) => startNotesDrag(e, 0)} />
                  )}
                  {showNotes && (
                    <div
                      className={`notes-pane${notesHeight < 60 ? ' notes-thin' : ''}`}
                      style={{ height: notesHeight }}
                    >
                      <div
                        className="notes-resize-handle"
                        onMouseDown={(e) => startNotesDrag(e, notesHeight)}
                      />
                      <textarea
                        value={notesText}
                        placeholder={
                          notesHeight < 60 ? t('appNotesClickToAdd') : t('appNotesPlaceholder')
                        }
                        onChange={(e) => onNotesChange(e.target.value)}
                        onBlur={() => void flushNotes()}
                      />
                    </div>
                  )}
                </div>
                {showBgFormat ? (
                  <FormatBackgroundPane
                    key={current}
                    slide={slide}
                    onApply={(op, all) => void onBackground(op, all)}
                    onCollapse={() => setShowBgFormat(false)}
                  />
                ) : showFormat ? (
                  <FormatPane
                    node={selectedNode}
                    onTransform={onTransform}
                    onFill={(id, fill) => void onFill(id, fill)}
                    onImageFill={(id) =>
                      void window.slidesApi
                        .editImageFill({ slideIndex: current, sourceId: id })
                        .then((r) => r && applySlide(current, r))
                    }
                    onTextAnchor={(id, anchor) =>
                      void window.slidesApi
                        .setTextAnchor({ slideIndex: current, sourceId: id, anchor })
                        .then((r) => r && applySlide(current, r))
                    }
                    onStroke={(id, stroke) => void onStroke(id, stroke)}
                    onDelete={() => void deleteSelected()}
                    onCollapse={() => setShowFormat(false)}
                    onPictureCrop={startCrop}
                    onPictureCutout={startCutout}
                    pictureCanCutout={contextPictureCanCutout}
                    link={selectedLink}
                    onOpenLink={openLinkDialog}
                    chartData={selectedChartData}
                    onChartPointColor={(si, pi, color) =>
                      void onEditChart({ pointColors: { [si]: { [pi]: color } } })
                    }
                  />
                ) : showAnimPane ? (
                  <AnimationPane
                    slideIndex={current}
                    items={animations}
                    selected={selAnim}
                    onSelect={setSelAnim}
                    onMove={moveAnimation}
                    onDelete={deleteAnimation}
                    onPreview={() => setAnimPreview((n) => n + 1)}
                    onCollapse={() => setShowAnimPane(false)}
                  />
                ) : showComments ? (
                  <CommentsPane
                    slideIndex={current}
                    comments={comments}
                    focusNonce={commentsFocusNonce}
                    onAdd={(text) => void addComment(text)}
                    onDelete={(c) => void deleteComment(c)}
                    onCollapse={() => setShowComments(false)}
                  />
                ) : null}
              </>
            )}
          </div>

          <footer className="status-bar">
            <div className="status-left">
              {slide ? (
                <span className="status-item">
                  {t('appStatusBarSlide', { current: current + 1, total: slides.length })}
                </span>
              ) : (
                t('appStatusBarReady')
              )}
              {status && <span className="status-msg"> — {status}</span>}
            </div>
            <div className="status-right">
              {hasDoc && (
                <button
                  className={`status-notes-btn${showNotes ? ' on' : ''}`}
                  data-tip={showNotes ? t('appNotesHide') : t('appNotesShow')}
                  onClick={() => setShowNotes((v) => !v)}
                >
                  <IconNotes size={18} />
                  <span>{t('appNotesLabel')}</span>
                </button>
              )}
              {hasDoc && (
                <button
                  className="status-play-btn"
                  data-tip={t('ribbonFromCurrentTip')}
                  aria-label={t('ribbonFromCurrentTip')}
                  onClick={() => startSlideShow(false)}
                >
                  <IconPlayBoxed size={18} />
                </button>
              )}
              <ZoomControls zoom={zoom} onPreview={previewZoom} />
            </div>
          </footer>
        </div>
      </div>

      {masterItems && (
        <MasterView
          key={masterItems[0]?.partPath}
          initialItems={masterItems}
          onClose={() => void closeMasterView()}
        />
      )}

      {slideShow && slides.length > 0 && (
        <SlideShowView
          slides={slides}
          images={images}
          startAt={slideShow.startAt}
          customOrder={slideShow.customOrder}
          rehearseMode={slideShow.rehearse}
          onRehearseDone={onRehearseDone}
          onExit={exitSlideShow}
        />
      )}

      {presenter && slides.length > 0 && (
        <PresenterView
          slides={slides}
          images={images}
          startAt={presenter.startAt}
          onExit={exitPresenterView}
          onUseSlideShow={switchPresenterToShow}
        />
      )}

      {chartDataDialogOpen && chartDataDialogInit && (
        <ChartDataDialog
          init={chartDataDialogInit}
          onClose={() => setChartDataDialogOpen(false)}
          onConfirm={(categories, series) => {
            setChartDataDialogOpen(false)
            void onEditChart({ categories, series })
          }}
        />
      )}

      {linkDialog && (
        <LinkDialog
          initial={linkDialog.initial}
          slideCount={slides.length}
          currentSlide={current}
          onApply={(target) => void applyLink(target)}
          onClose={() => setLinkDialog(null)}
        />
      )}
      {hfDialog && (
        <HeaderFooterDialog
          initial={hfDialog}
          onApply={(opts) => void applyHf(opts)}
          onClose={() => setHfDialog(null)}
        />
      )}
      {eqDialogOpen && (
        <EquationDialog
          onInsert={(text) => void insertEquation(text)}
          onClose={() => setEqDialogOpen(false)}
        />
      )}

      {cutoutTarget && (
        <CutoutDialog
          dataUrl={cutoutTarget.dataUrl}
          onApply={(png) => void applyCutout(png)}
          onCancel={() => setCutoutTarget(null)}
        />
      )}

      {customShowDlgOpen && (
        <CustomShowDialog
          shows={customShows}
          slideCount={slides.length}
          onChange={updateCustomShows}
          onPlay={playCustomShow}
          onClose={() => setCustomShowDlgOpen(false)}
        />
      )}

      {printDlgOpen && (
        <PrintDialog
          slides={slides}
          images={images}
          current={current}
          onClose={() => setPrintDlgOpen(false)}
          setStatus={setStatus}
        />
      )}

      {findOpen && (
        <FindReplaceDialog
          slides={slides}
          onNavigate={(si, id) => {
            setCurrent(si)
            setSelectedIds([id])
          }}
          onReplaced={(all) => {
            setSlides(all)
            setDirty(true)
          }}
          onClose={() => setFindOpen(false)}
        />
      )}

      {pendingRehearse && (
        <div className="modal-backdrop">
          <div className="modal">
            <h2>{t('appRehearseTitle')}</h2>
            <p className="rehearse-summary">
              {t('appRehearseSummary', {
                duration: formatClock(pendingRehearse.reduce((a, b) => a + b, 0) * 1000),
              })}
            </p>
            <div className="modal-actions">
              <button onClick={() => setPendingRehearse(null)}>{t('appRehearseDiscard')}</button>
              <button className="primary" onClick={() => void saveRehearseTimings()}>
                {t('appRehearseSave')}
              </button>
            </div>
          </div>
        </div>
      )}

      {ctxMenu && (
        <ContextMenu
          x={ctxMenu.x}
          y={ctxMenu.y}
          items={ctxItems}
          onClose={() => setCtxMenu(null)}
        />
      )}
    </div>
  )
}

/** Clamp a zoom-anchor coordinate into the viewport span; undefined passes through (→ center). */
function anchorClamp(v: number | undefined, min: number, max: number): number | undefined {
  return v == null ? undefined : Math.min(Math.max(v, min), max)
}

/** Status-bar zoom controls. The slider keeps a local live value so the thumb and the %
 * label track mid-drag while only this tiny component re-renders per tick; the drag goes
 * through the shared preview path (CSS transform now, debounced App-level commit at the
 * gesture's end) — a setZoom per tick re-renders the whole App and stutters. */
function ZoomControls({
  zoom,
  onPreview,
}: {
  readonly zoom: number
  readonly onPreview: (z: number | ((current: number) => number)) => void
}) {
  const [live, setLive] = useState(() => Math.round(zoom * 100))
  // adopt outside commits (fit, pinch, menu) once they land
  useEffect(() => setLive(Math.round(zoom * 100)), [zoom])
  // Buttons go through the preview path too: the zoom pivots on the viewport center and
  // rapid clicks compound on the pending value; `live` mirrors it so the % keeps up
  const step = (dir: 1 | -1) => {
    setLive((v) => Math.min(300, Math.max(25, v + dir * 10)))
    onPreview((z) => Math.min(3, Math.max(0.25, z + dir * 0.1)))
  }
  return (
    <>
      <button className="zoom-btn" onClick={() => step(-1)}>
        −
      </button>
      <input
        className="zoom-slider"
        type="range"
        min={25}
        max={300}
        step={5}
        value={live}
        onChange={(e) => {
          const v = Number(e.target.value)
          setLive(v)
          onPreview(v / 100)
        }}
      />
      <button className="zoom-btn" onClick={() => step(1)}>
        +
      </button>
      <span className="zoom-value">{live}%</span>
    </>
  )
}
