import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, ReactElement, ReactNode, RefObject } from 'react'
// legacy build: the modern build relies on new APIs like Math.sumPrecise that the current
// Electron V8 lacks, making embedded font parsing fail and whole pages render as garbled raw char codes
import { GlobalWorkerOptions, TextLayer, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { AiPanel, GensparkMark } from './ai/AiPanel'
import type { PdfAiDeps } from './ai/tools'
import {
  MARKUP_COLORS,
  geomDispSize,
  pdfRectToCss,
  quadToRect,
  selectionQuadsByPage,
  viewToPdf,
} from './annotations'
import type { LocalMarkup, PageGeom } from './annotations'
import { DRAW_COLORS, DrawLayer, cssRgb } from './DrawLayer'
import type { DrawTool, LocalDrawing } from './DrawLayer'
import { FormLayer } from './FormLayer'
import { LinkLayer } from './LinkLayer'
import { OutlinePanel } from './OutlinePanel'
import type { OutlineNode } from './OutlinePanel'
import { printPdf } from './print'
import { PropertiesDialog } from './PropertiesDialog'
import { SignatureDialog } from './SignatureDialog'
import type { SignatureStrokes } from './SignatureDialog'
import { StampDialog } from './StampDialog'
import { buildStamps } from './stamps'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import { buildSearchIndex, searchInIndex } from './search'
import type { SearchIndex, SearchMatch } from './search'
import { useI18n } from './i18n/locale'
import { useAutosave } from './useAutosave'
import type {
  DrawingInput,
  FormValueInput,
  MarkupType,
  MetadataInput,
  StampInput,
} from '../shared/ipc'

GlobalWorkerOptions.workerSrc = workerUrl

// cmaps/standard fonts/wasm are statically copied by the build into pdfjs/ of the renderer output (same path on the dev server)
const ASSET_BASE = new URL('pdfjs/', document.baseURI).href

const ZOOM_STEPS = [0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]
const MIN_SCALE = ZOOM_STEPS[0]
const MAX_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]
const PAGE_GAP = 16
const SCROLL_PAD = 24
const THUMB_WIDTH = 108

interface PageSize {
  width: number
  height: number
}

type FitMode = 'width' | 'page' | null

const DOC_OPTS = {
  cMapUrl: `${ASSET_BASE}cmaps/`,
  cMapPacked: true,
  standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
  wasmUrl: `${ASSET_BASE}wasm/`,
}

/** Which items in the container are within the (expanded) viewport — shared lazy-render basis
    for pages/thumbnails. Rebuild the observer when enabled flips (sidebar toggles unmount/remount the root) */
function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
  enabled = true,
): { visible: Set<number>; setItemRef: (idx: number) => (el: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])
  useEffect(() => {
    const root = rootRef.current
    if (!enabled || !root || count === 0) return
    const io = new IntersectionObserver(
      (entries) => {
        setVisible((prev) => {
          const next = new Set(prev)
          for (const e of entries) {
            const idx = Number((e.target as HTMLElement).dataset.idx)
            if (e.isIntersecting) next.add(idx)
            else next.delete(idx)
          }
          return next
        })
      },
      { root, rootMargin },
    )
    for (const el of itemRefs.current) if (el) io.observe(el)
    return () => io.disconnect()
  }, [rootRef, count, rootMargin, enabled])
  return {
    visible,
    setItemRef: (idx) => (el) => {
      itemRefs.current[idx] = el
    },
  }
}

/** Single page: renders canvas + text layer (select/copy) when visible, released once off-viewport */
function PdfPage({
  doc,
  pageNo,
  scale,
  rotationDelta,
  visible,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  /** Unsaved rotation delta (clockwise degrees) */
  rotationDelta: number
  visible: boolean
}) {
  const holderRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const holder = holderRef.current
    if (!visible || !holder) return
    let cancelled = false
    let renderTask: RenderTask | null = null
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const viewport = page.getViewport({ scale, rotation: (page.rotate + rotationDelta) % 360 })
      // Cap at 2x: on hi-dpi screens a 3x-dpr full-page bitmap doubles memory with no visible gain
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      const canvas = document.createElement('canvas')
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        return // cancelled
      }
      if (cancelled) return
      const textDiv = document.createElement('div')
      textDiv.className = 'textLayer'
      holder.replaceChildren(canvas, textDiv)
      const textLayer = new TextLayer({
        textContentSource: page.streamTextContent(),
        container: textDiv,
        viewport,
      })
      try {
        await textLayer.render()
      } catch {
        /* cancelled */
      }
    })()
    return () => {
      cancelled = true
      renderTask?.cancel()
      holder.replaceChildren()
    }
  }, [doc, pageNo, scale, rotationDelta, visible])
  return <div ref={holderRef} className="pdf-page-content" />
}

/** Overlay for unsaved markups; click to select (deletion is explicit via the delete popup or Delete key) */
function MarkupOverlay({
  markups,
  geom,
  scale,
  selectedId,
  selectTitle,
  onSelect,
}: {
  markups: LocalMarkup[]
  geom: PageGeom
  scale: number
  selectedId: string | null
  selectTitle: string
  onSelect: (id: string, x: number, y: number) => void
}) {
  return (
    <>
      {markups.flatMap((m) =>
        m.quads.map((q, i) => {
          const [r, g, b] = m.color
          const style: CSSProperties = pdfRectToCss(geom, quadToRect(q), scale)
          if (m.type === 'highlight') {
            style.background = `rgba(${r * 255}, ${g * 255}, ${b * 255}, 0.4)`
          } else {
            const bar = `rgb(${r * 255}, ${g * 255}, ${b * 255})`
            if (m.type === 'underline') style.borderBottom = `2px solid ${bar}`
            else style.backgroundImage = `linear-gradient(${bar}, ${bar})`
          }
          return (
            <div
              key={`${m.id}-${i}`}
              className={`pdf-markup pdf-markup-${m.type}${m.id === selectedId ? ' pdf-markup-selected' : ''}`}
              style={style}
              title={selectTitle}
              onClick={(e) => onSelect(m.id, e.clientX, e.clientY)}
            />
          )
        }),
      )}
    </>
  )
}

/** Thumbnail: rendered once per (doc, rotation) when visible and cached */
function PdfThumb({
  doc,
  pageNo,
  rotationDelta,
  visible,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  rotationDelta: number
  visible: boolean
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const renderedKeyRef = useRef<string | null>(null)
  useEffect(() => {
    const canvas = canvasRef.current
    const key = `${rotationDelta}`
    if (!visible || !canvas || renderedKeyRef.current === key) return
    renderedKeyRef.current = key
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const rotation = (page.rotate + rotationDelta) % 360
      const scale = THUMB_WIDTH / page.getViewport({ scale: 1, rotation }).width
      const viewport = page.getViewport({ scale, rotation })
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      try {
        await page.render({
          canvas,
          viewport,
          transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
        }).promise
      } catch {
        renderedKeyRef.current = null
      }
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo, rotationDelta, visible])
  // Reset the cache key when the doc changes (save reload); re-render next time it's visible
  useEffect(() => {
    renderedKeyRef.current = null
  }, [doc])
  return <canvas ref={canvasRef} style={{ width: THUMB_WIDTH }} />
}

// ── ribbon icons (aligned with slides' rb-big visual language) ──

function Icon({ children }: { children: ReactNode }): ReactElement {
  return (
    <svg
      width="20"
      height="20"
      viewBox="0 0 20 20"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.4"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

const IconThumbs = () => (
  <Icon>
    <rect x="2.5" y="3" width="6" height="6.5" rx="1" />
    <rect x="2.5" y="12" width="6" height="5" rx="1" />
    <path d="M12 4h5.5M12 8h5.5M12 13h5.5M12 16h5.5" />
  </Icon>
)
const IconHighlight = () => (
  <Icon>
    <path d="M4 16.5 12.5 8l3.5 3.5-8.5 8.5H4v-3.5Z" transform="translate(0 -3)" />
    <path d="M12.5 5 15 2.5 18.5 6 16 8.5" transform="translate(0 -1)" />
    <path d="M3 18.5h14" stroke="#ffc21c" strokeWidth="2.4" />
  </Icon>
)
const IconUnderline = () => (
  <Icon>
    <path d="M6 3.5v6a4 4 0 0 0 8 0v-6" />
    <path d="M4.5 17h11" stroke="#2b66ff" strokeWidth="2" />
  </Icon>
)
const IconStrike = () => (
  <Icon>
    <path d="M13.8 6c-.5-1.5-2-2.4-3.8-2.4-2.2 0-3.8 1.2-3.8 3 0 1.4.9 2.2 2.6 2.7" />
    <path d="M7 14.6c.5 1.4 2 2.3 3.9 2.3 2.2 0 3.9-1.1 3.9-3 0-.6-.1-1.1-.4-1.5" />
    <path d="M3.5 10.5h13" stroke="#db3830" strokeWidth="1.8" />
  </Icon>
)
const IconInk = () => (
  <Icon>
    <path d="M3 15.5c2.5.5 3.5-1 3.5-3s-1.2-3.4-2.6-2.6c-1.2.7-.6 3 1.6 4.3 2.4 1.4 5.2.6 7.4-1.6 1.6-1.6 2.6-3.6 3.1-5.6" />
    <path d="M14.5 3.5 17 6l-1.5 1.5" />
  </Icon>
)
const IconRect = () => (
  <Icon>
    <rect x="3" y="5" width="14" height="10" rx="1" />
  </Icon>
)
const IconEllipse = () => (
  <Icon>
    <ellipse cx="10" cy="10" rx="7" ry="5.2" />
  </Icon>
)
const IconArrow = () => (
  <Icon>
    <path d="M3.5 15.5 16 3.5" />
    <path d="M10.5 3.5H16v5.5" />
  </Icon>
)
const IconNote = () => (
  <Icon>
    <path d="M3 4.5h14v9h-8l-4 3v-3H3z" strokeLinejoin="round" />
    <path d="M6.5 8h7M6.5 10.5h4.5" />
  </Icon>
)
const IconSign = () => (
  <Icon>
    <path d="M2.5 14.5c2-.3 3.4-1.6 4.6-3.6 1.1-1.9 1.7-4 1.3-5.4-.3-1-1.2-1-1.6 0-.5 1.3-.3 3.6.7 5.6 1 2 2.4 3.1 3.9 3.1 1.2 0 2.1-.6 2.6-1.6" />
    <path d="M3 17.5h14" />
  </Icon>
)
const IconExportImg = () => (
  <Icon>
    <rect x="2.5" y="4" width="15" height="10.5" rx="1" />
    <circle cx="7" cy="7.8" r="1.2" />
    <path d="M2.8 12.2 7 9l3.4 2.6L13 9.6l4.2 3.4" />
  </Icon>
)
const IconNight = () => (
  <Icon>
    <path d="M15.5 12.2A6.5 6.5 0 0 1 7.8 4.5a6.5 6.5 0 1 0 7.7 7.7Z" />
  </Icon>
)
const IconSpread = () => (
  <Icon>
    <rect x="2.5" y="4" width="6.5" height="12" rx="1" />
    <rect x="11" y="4" width="6.5" height="12" rx="1" />
  </Icon>
)
const IconSinglePage = () => (
  <Icon>
    <rect x="5.5" y="3.5" width="9" height="13" rx="1" />
  </Icon>
)
const IconWatermark = () => (
  <Icon>
    <rect x="3" y="3.5" width="14" height="13" rx="1" />
    <path d="M6 13.5 13.5 6" />
    <path d="M6 9.5 9.5 6M10.5 13.5 14 10" />
  </Icon>
)
const IconProps = () => (
  <Icon>
    <circle cx="10" cy="10" r="7.2" />
    <path d="M10 9v4.5" />
    <circle cx="10" cy="6.6" r="0.9" fill="currentColor" stroke="none" />
  </Icon>
)
const IconRotateL = () => (
  <Icon>
    <path d="M6.5 7.5H3V4" />
    <path d="M3.2 7.2A7 7 0 1 1 3 10" />
  </Icon>
)
const IconRotateR = () => (
  <Icon>
    <path d="M13.5 7.5H17V4" />
    <path d="M16.8 7.2A7 7 0 1 0 17 10" />
  </Icon>
)
const IconDeletePage = () => (
  <Icon>
    <path d="M5 3h7l3 3v11a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z" />
    <path d="M7.5 11l5 5M12.5 11l-5 5" stroke="#db3830" />
  </Icon>
)
const IconExtract = () => (
  <Icon>
    <path d="M5 3h7l3 3v5" />
    <path d="M4 3.9V16a1 1 0 0 0 1 1h5" />
    <path d="M15 13.5v6M12 16.5l3 3 3-3" transform="translate(0 -2)" />
  </Icon>
)
const IconInsertPdf = () => (
  <Icon>
    <path d="M5 3h7l3 3v11a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z" />
    <path d="M10 8.5v6M7 11.5h6" stroke="#217346" />
  </Icon>
)
const IconFitWidth = () => (
  <Icon>
    <path d="M3 4v12M17 4v12" />
    <path d="M6 10h8M8.2 7.8 6 10l2.2 2.2M11.8 7.8 14 10l-2.2 2.2" />
  </Icon>
)
const IconFitPage = () => (
  <Icon>
    <rect x="5" y="2.5" width="10" height="15" rx="1" />
    <path d="M10 6v8M7.8 8.2 10 6l2.2 2.2M7.8 11.8 10 14l2.2-2.2" />
  </Icon>
)
const IconOutline = () => (
  <Icon>
    <path d="M3.5 4.5h13M6.5 8.5h10M6.5 12.5h10M9.5 16.5h7" />
    <circle cx="4" cy="8.5" r="0.8" fill="currentColor" />
    <circle cx="4" cy="12.5" r="0.8" fill="currentColor" />
    <circle cx="7" cy="16.5" r="0.8" fill="currentColor" />
  </Icon>
)
const IconUndo = () => (
  <svg
    width="15"
    height="15"
    viewBox="0 0 16 16"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.3"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M5.5 3.5 2.5 6.5l3 3" />
    <path d="M2.5 6.5h7a4 4 0 0 1 0 8H6" />
  </svg>
)
const IconRedo = () => (
  <svg
    width="15"
    height="15"
    viewBox="0 0 16 16"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.3"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M10.5 3.5 13.5 6.5l-3 3" />
    <path d="M13.5 6.5h-7a4 4 0 0 0 0 8H10" />
  </svg>
)
const IconSearch = () => (
  <Icon>
    <circle cx="9" cy="9" r="5.5" />
    <path d="M13.2 13.2 17 17" />
  </Icon>
)
const IconPrint = () => (
  <Icon>
    <path d="M6 7V3h8v4" />
    <rect x="3.5" y="7" width="13" height="6.5" rx="1" />
    <path d="M6 11h8v6H6z" />
  </Icon>
)
const IconSave = () => (
  <svg
    width="15"
    height="15"
    viewBox="0 0 16 16"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M2.5 3.5a1 1 0 0 1 1-1h7.6l2.4 2.4v7.6a1 1 0 0 1-1 1h-9a1 1 0 0 1-1-1v-9Z" />
    <path d="M5 2.7v3.1h5.4V2.7M4.8 13.3V9h6.4v4.3" />
  </svg>
)

interface ThumbMenu {
  x: number
  y: number
  origIdx: number
}

const DRAW_TOOLS = [
  { tool: 'ink' as const, icon: IconInk, key: 'drawInk' as const },
  { tool: 'rect' as const, icon: IconRect, key: 'drawRect' as const },
  { tool: 'ellipse' as const, icon: IconEllipse, key: 'drawEllipse' as const },
  { tool: 'arrow' as const, icon: IconArrow, key: 'drawArrow' as const },
  { tool: 'note' as const, icon: IconNote, key: 'drawNote' as const },
]

/** Drawing stroke width (PDF pt); thin lines stay crisp under zoom */
const STROKE_WIDTH = 2

/** Full snapshot of unsaved edits (for undo/redo; data is small, whole-copy replace is safest) */
/** Watermark/header-footer are kept as config and rendered in final page order only at save time, so page numbers survive reorders/deletions */
interface StampConfig {
  wm: WatermarkConfig | null
  hf: HeaderFooterConfig | null
}

interface EditSnapshot {
  markups: LocalMarkup[]
  drawings: LocalDrawing[]
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  deleted: Set<number>
  order: number[] | null
  metadata: MetadataInput | null
}

/** Selected annotation with the anchor of its floating delete popup; a stamp click selects the whole watermark/header-footer set */
type AnnotSelection =
  | { kind: 'markup' | 'drawing'; id: string; x: number; y: number }
  | { kind: 'stamp'; x: number; y: number }

/** Page ranges like "1-3,5" → list of 1-based page numbers; null if invalid */
function parsePageRanges(input: string, max: number): number[] | null {
  const out = new Set<number>()
  for (const part of input.split(/[,，]/)) {
    const s = part.trim()
    if (!s) continue
    const m = /^(\d+)\s*[-–]\s*(\d+)$|^(\d+)$/.exec(s)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i)
  }
  return out.size > 0 ? [...out].sort((x, y) => x - y) : null
}

export default function App() {
  const { t } = useI18n()
  const [doc, setDoc] = useState<PDFDocumentProxy | null>(null)
  const [filePath, setFilePath] = useState('')
  const [status, setStatus] = useState<'loading' | 'error' | 'empty' | 'password' | 'ready'>(
    'loading',
  )
  const [sizes, setSizes] = useState<PageSize[]>([])
  const [baseRots, setBaseRots] = useState<number[]>([])
  const [scale, setScale] = useState(1)
  const [currentPage, setCurrentPage] = useState(1)
  const [pageInput, setPageInput] = useState('1')
  const [sidebar, setSidebar] = useState<'thumbs' | 'outline' | null>('thumbs')
  const [aiCollapsed, setAiCollapsed] = useState(false)
  const [spread, setSpread] = useState<1 | 2>(1)
  const [nightMode, setNightMode] = useState(false)
  const [outline, setOutline] = useState<OutlineNode[] | null>(null)
  const [markups, setMarkups] = useState<LocalMarkup[]>([])
  const [drawings, setDrawings] = useState<LocalDrawing[]>([])
  const [drawTool, setDrawTool] = useState<DrawTool | null>(null)
  const [drawColor, setDrawColor] = useState<[number, number, number]>(DRAW_COLORS[0]!.rgb)
  const [notePrompt, setNotePrompt] = useState<{ origIdx: number; at: [number, number] } | null>(
    null,
  )
  const [noteText, setNoteText] = useState('')
  const [stampCfg, setStampCfg] = useState<StampConfig | null>(null)
  /** User-defined page order (original page indices); null means unreordered */
  const [order, setOrder] = useState<number[] | null>(null)
  const [metadata, setMetadata] = useState<MetadataInput | null>(null)
  const [stampDlg, setStampDlg] = useState(false)
  const [propsDlg, setPropsDlg] = useState(false)
  const [fileSize, setFileSize] = useState(0)
  const [dragFrom, setDragFrom] = useState<number | null>(null)
  const [dragOver, setDragOver] = useState<number | null>(null)
  const [signDlg, setSignDlg] = useState(false)
  /** Confirmed signature awaiting placement; when non-null the page enters click-to-place mode */
  const [pendingSign, setPendingSign] = useState<SignatureStrokes | null>(null)
  const [exporting, setExporting] = useState(false)
  const [formEdits, setFormEdits] = useState<Map<string, FormValueInput>>(new Map())
  const [rotations, setRotations] = useState<Map<number, number>>(new Map())
  const [deleted, setDeleted] = useState<Set<number>>(new Set())
  const [selPopup, setSelPopup] = useState<{ x: number; y: number } | null>(null)
  const [selected, setSelected] = useState<AnnotSelection | null>(null)
  const [deleteToast, setDeleteToast] = useState(false)
  const toastTimerRef = useRef<number | null>(null)
  const [thumbMenu, setThumbMenu] = useState<ThumbMenu | null>(null)
  const [saveState, setSaveState] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [saveError, setSaveError] = useState('')
  /** Autosave gate: this file was saved explicitly at least once */
  const savedOnceRef = useRef(false)
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchMatches, setSearchMatches] = useState<SearchMatch[]>([])
  const [searchCur, setSearchCur] = useState(0)
  const [printing, setPrinting] = useState(false)
  const [undoStack, setUndoStack] = useState<EditSnapshot[]>([])
  const [redoStack, setRedoStack] = useState<EditSnapshot[]>([])
  const [pwInput, setPwInput] = useState('')
  const [pwWrong, setPwWrong] = useState(false)
  const [extractDlg, setExtractDlg] = useState(false)
  const [extractInput, setExtractInput] = useState('')
  const [extractInvalid, setExtractInvalid] = useState(false)
  const coalesceKeyRef = useRef<string | null>(null)
  const passwordRef = useRef<string | undefined>(undefined)
  const fitModeRef = useRef<FitMode>('width')
  const scrollRef = useRef<HTMLDivElement>(null)
  const thumbsRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const searchIndexRef = useRef<{ doc: PDFDocumentProxy; promise: Promise<SearchIndex> } | null>(
    null,
  )
  const searchJumpRef = useRef<{ matches: SearchMatch[]; cur: number } | null>(null)

  /** Visible pages (with unsaved reorder, deleted pages hidden): position → original page index */
  const visList = useMemo(() => {
    const base = order ?? sizes.map((_, i) => i)
    return base.filter((i) => !deleted.has(i))
  }, [sizes, deleted, order])
  const pageCount = visList.length

  /** Render rows: one page per row in single mode, two in spread mode (first page alone, like a book cover) */
  const rows = useMemo(() => {
    if (spread === 1) return visList.map((i) => [i])
    const out: number[][] = []
    for (let i = 0; i < visList.length; i++) {
      if (i === 0) out.push([visList[0]!])
      else if (i % 2 === 1)
        out.push([visList[i]!, ...(visList[i + 1] !== undefined ? [visList[i + 1]!] : [])])
    }
    return out
  }, [visList, spread])

  /** Visible position → row index */
  const rowOfVis = useCallback(
    (visIdx: number) =>
      spread === 1 ? visIdx : visIdx === 0 ? 0 : Math.floor((visIdx - 1) / 2) + 1,
    [spread],
  )
  const fileName = filePath.split(/[\\/]/).pop() ?? filePath

  const rotDelta = useCallback((origIdx: number) => rotations.get(origIdx) ?? 0, [rotations])
  /** Page geometry: unrotated size + total display rotation; the single entry point for overlay coord conversion */
  const pageGeom = useCallback(
    (origIdx: number): PageGeom => {
      const s = sizes[origIdx]!
      return { pw: s.width, ph: s.height, rot: (baseRots[origIdx] ?? 0) + rotDelta(origIdx) }
    },
    [sizes, baseRots, rotDelta],
  )
  /** Page display size (width/height swapped under rotation) */
  const dispSize = useCallback(
    (origIdx: number): PageSize => geomDispSize(pageGeom(origIdx)),
    [pageGeom],
  )

  const { visible: visibleRows, setItemRef: setRowRef } = useVisibleSet(
    scrollRef,
    rows.length,
    '800px 0px',
  )
  const { visible: visibleThumbs, setItemRef: setThumbRef } = useVisibleSet(
    thumbsRef,
    pageCount,
    '400px 0px',
    sidebar === 'thumbs',
  )

  const loadDoc = useCallback(async (path: string, previous: PDFDocumentProxy | null) => {
    const data = await window.pdfApi.readFile(path)
    const loaded = await getDocument({
      data: new Uint8Array(data),
      password: passwordRef.current,
      ...DOC_OPTS,
    }).promise
    const all: PageSize[] = []
    const rots: number[] = []
    for (let i = 1; i <= loaded.numPages; i++) {
      const page = await loaded.getPage(i)
      // Unrotated size; display size is derived by geom from the total rotation
      const vp = page.getViewport({ scale: 1, rotation: 0 })
      all.push({ width: vp.width, height: vp.height })
      rots.push(page.rotate ?? 0)
    }
    setSizes(all)
    setBaseRots(rots)
    setDoc(loaded)
    setMarkups([])
    setDrawings([])
    setStampCfg(null)
    setFormEdits(new Map())
    setRotations(new Map())
    setDeleted(new Set())
    setOrder(null)
    setMetadata(null)
    setFileSize(data.byteLength)
    setSelected(null)
    setDeleteToast(false)
    setUndoStack([])
    setRedoStack([])
    void loaded.getOutline().then(
      (o) => setOutline(o && o.length > 0 ? (o as OutlineNode[]) : null),
      () => setOutline(null),
    )
    if (previous) void previous.destroy()
  }, [])

  const openPath = useCallback(
    async (path: string) => {
      try {
        setFilePath(path)
        // A newly opened file starts outside the autosave gate
        savedOnceRef.current = false
        await loadDoc(path, null)
        setStatus('ready')
      } catch (err) {
        if ((err as Error | null)?.name === 'PasswordException') {
          setPwWrong(passwordRef.current !== undefined)
          setStatus('password')
          return
        }
        console.error('[pdf] open failed:', err)
        setStatus('error')
      }
    },
    [loadDoc],
  )

  useEffect(() => {
    void (async () => {
      const path = await window.pdfApi.consumePending()
      if (!path) {
        setStatus('empty')
        return
      }
      await openPath(path)
    })()
  }, [openPath])

  /** Documents opened with a password are treated as read-only: pdf-lib can't write back encrypted files */
  const readOnly = status === 'ready' && passwordRef.current !== undefined

  const clampScale = (s: number) => Math.min(MAX_SCALE, Math.max(MIN_SCALE, s))

  /** Overall size of a row (in spread mode widths add up, including the page gap) */
  const rowSize = useCallback(
    (row: number[]): PageSize => {
      const dims = row.map((i) => dispSize(i))
      return {
        width: dims.reduce((w, d) => w + d.width, 0) + (dims.length - 1) * PAGE_GAP,
        height: Math.max(...dims.map((d) => d.height)),
      }
    },
    [dispSize],
  )

  const recomputeFit = useCallback(() => {
    const mode = fitModeRef.current
    const el = scrollRef.current
    if (!mode || !el || rows.length === 0) return
    const dims = rows.map((r) => rowSize(r))
    const maxW = Math.max(...dims.map((s) => s.width))
    const availW = el.clientWidth - SCROLL_PAD * 2
    if (mode === 'width') {
      setScale(clampScale(availW / maxW))
    } else {
      const maxH = Math.max(...dims.map((s) => s.height))
      setScale(clampScale(Math.min(availW / maxW, (el.clientHeight - PAGE_GAP * 2) / maxH)))
    }
  }, [rows, rowSize])

  useEffect(() => {
    if (status !== 'ready') return
    recomputeFit()
    const el = scrollRef.current
    if (!el) return
    const ro = new ResizeObserver(recomputeFit)
    ro.observe(el)
    return () => ro.disconnect()
  }, [status, recomputeFit])

  /** Cumulative row-top offset (with gaps), shared by scroll positioning and current-page calc */
  const rowTop = useCallback(
    (rowIdx: number) => {
      let y = PAGE_GAP
      for (let i = 0; i < rowIdx; i++) y += rowSize(rows[i]!).height * scale + PAGE_GAP
      return y
    },
    [rows, rowSize, scale],
  )

  /** Page-top offset of a visible position (used for search positioning) */
  const pageTop = useCallback((visIdx: number) => rowTop(rowOfVis(visIdx)), [rowTop, rowOfVis])

  const handleScroll = useCallback(() => {
    const el = scrollRef.current
    if (!el || rows.length === 0) return
    const anchor = el.scrollTop + el.clientHeight * 0.4
    let rowIdx = 0
    for (let i = 0; i < rows.length; i++) {
      if (rowTop(i) <= anchor) rowIdx = i
      else break
    }
    const page = visList.indexOf(rows[rowIdx]![0]!) + 1
    setCurrentPage(page)
    setPageInput(String(page))
  }, [rows, rowTop, visList])

  const scrollToPage = (n: number) => {
    const el = scrollRef.current
    if (!el) return
    const target = Math.min(Math.max(1, n), pageCount)
    el.scrollTop = rowTop(rowOfVis(target - 1)) - PAGE_GAP / 2
  }

  /** Scale scroll position proportionally when zooming so the visual anchor doesn't jump */
  const applyScale = (next: number, mode: FitMode) => {
    fitModeRef.current = mode
    const el = scrollRef.current
    const clamped = clampScale(next)
    if (el && scale > 0) {
      const ratio = clamped / scale
      requestAnimationFrame(() => {
        el.scrollTop *= ratio
      })
    }
    setScale(clamped)
  }

  const zoomIn = () => applyScale(ZOOM_STEPS.find((s) => s > scale + 0.001) ?? MAX_SCALE, null)
  const zoomOut = () =>
    applyScale([...ZOOM_STEPS].reverse().find((s) => s < scale - 0.001) ?? MIN_SCALE, null)

  const commitPageInput = () => {
    const n = Number.parseInt(pageInput, 10)
    if (Number.isFinite(n)) scrollToPage(n)
    else setPageInput(String(currentPage))
  }

  const dirty =
    markups.length > 0 ||
    drawings.length > 0 ||
    stampCfg !== null ||
    formEdits.size > 0 ||
    rotations.size > 0 ||
    deleted.size > 0 ||
    order !== null ||
    metadata !== null

  // Mirror dirty state to the main process (close-tab/close-window guard)
  useEffect(() => {
    window.pdfApi.setDirty(dirty)
  }, [dirty])

  // ── Undo/redo: push a full snapshot before each change; consecutive input on the same form field coalesces into one step ──

  const snapshot = (): EditSnapshot => ({
    markups,
    drawings,
    stampCfg,
    formEdits,
    rotations,
    deleted,
    order,
    metadata,
  })

  const pushUndo = (coalesceKey?: string) => {
    if (coalesceKey && coalesceKeyRef.current === coalesceKey) return
    coalesceKeyRef.current = coalesceKey ?? null
    setUndoStack((prev) => [...prev.slice(-49), snapshot()])
    setRedoStack([])
  }

  const applySnapshot = (s: EditSnapshot) => {
    setMarkups(s.markups)
    setDrawings(s.drawings)
    setStampCfg(s.stampCfg)
    setFormEdits(s.formEdits)
    setRotations(s.rotations)
    setDeleted(s.deleted)
    setOrder(s.order)
    setMetadata(s.metadata)
    // The selected annotation may no longer exist in the restored snapshot
    setSelected(null)
  }

  const undo = () => {
    const top = undoStack[undoStack.length - 1]
    if (!top) return
    setRedoStack((r) => [...r, snapshot()])
    setUndoStack((u) => u.slice(0, -1))
    applySnapshot(top)
    coalesceKeyRef.current = null
  }

  const redo = () => {
    const top = redoStack[redoStack.length - 1]
    if (!top) return
    setUndoStack((u) => [...u, snapshot()])
    setRedoStack((r) => r.slice(0, -1))
    applySnapshot(top)
    coalesceKeyRef.current = null
  }

  // ── Full-text search ──

  /** Text index cached per doc; invalidated and rebuilt after a save reload */
  const getSearchIndex = useCallback((): Promise<SearchIndex> | null => {
    if (!doc) return null
    if (searchIndexRef.current?.doc !== doc) {
      searchIndexRef.current = { doc, promise: buildSearchIndex(doc) }
    }
    return searchIndexRef.current.promise
  }, [doc])

  useEffect(() => {
    if (!searchOpen || !searchQuery.trim()) {
      setSearchMatches([])
      setSearchCur(0)
      return
    }
    let cancelled = false
    const timer = setTimeout(() => {
      void getSearchIndex()?.then((idx) => {
        if (cancelled) return
        setSearchMatches(searchInIndex(idx, searchQuery.trim()))
        setSearchCur(0)
      })
    }, 200)
    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [searchOpen, searchQuery, getSearchIndex])

  /** Pages with unsaved deletion are excluded from match navigation */
  const activeMatches = useMemo(
    () => searchMatches.filter((m) => !deleted.has(m.pageIndex)),
    [searchMatches, deleted],
  )
  const searchCurClamped = Math.min(searchCur, Math.max(0, activeMatches.length - 1))

  const gotoMatch = useCallback(
    (idx: number) => {
      const m = activeMatches[idx]
      const el = scrollRef.current
      if (!m || !el) return
      const visIdx = visList.indexOf(m.pageIndex)
      if (visIdx < 0) return
      const box = pdfRectToCss(pageGeom(m.pageIndex), m.rects[0] ?? [0, 0, 0, 0], scale)
      el.scrollTop = Math.max(0, pageTop(visIdx) + box.top - el.clientHeight * 0.35)
    },
    [activeMatches, visList, pageGeom, scale, pageTop],
  )

  // Scroll to the current match on new results or position changes (unrelated changes like zoom don't re-scroll)
  useEffect(() => {
    if (!searchOpen || activeMatches.length === 0) return
    const last = searchJumpRef.current
    if (last && last.matches === activeMatches && last.cur === searchCurClamped) return
    searchJumpRef.current = { matches: activeMatches, cur: searchCurClamped }
    gotoMatch(searchCurClamped)
  }, [searchOpen, activeMatches, searchCurClamped, gotoMatch])

  const searchStep = (dir: 1 | -1) => {
    const n = activeMatches.length
    if (n === 0) return
    setSearchCur((searchCurClamped + dir + n) % n)
  }

  const openSearch = () => {
    setSearchOpen(true)
    requestAnimationFrame(() => {
      searchInputRef.current?.focus()
      searchInputRef.current?.select()
    })
  }

  const closeSearch = () => setSearchOpen(false)

  /** Mouse released over selected text → show the markup bar centered above the selection box (below if it doesn't fit) */
  const handleMouseUp = () => {
    setTimeout(() => {
      const el = scrollRef.current
      const sel = window.getSelection()
      if (!el || !sel || sel.isCollapsed || sel.rangeCount === 0) {
        setSelPopup(null)
        return
      }
      if (!el.contains(sel.getRangeAt(0).commonAncestorContainer)) return
      if (readOnly) return
      const box = sel.getRangeAt(0).getBoundingClientRect()
      if (box.width < 1 && box.height < 1) return
      setSelPopup({
        x: Math.min(Math.max(box.left + box.width / 2, 70), window.innerWidth - 70),
        y: box.top >= 52 ? box.top - 44 : Math.min(box.bottom + 8, window.innerHeight - 44),
      })
    }, 0)
  }

  const applyMarkup = (type: MarkupType) => {
    const el = scrollRef.current
    if (!el || readOnly) return
    const byVisPage = selectionQuadsByPage(
      el,
      visList.map((i) => pageGeom(i)),
      scale,
    )
    setSelPopup(null)
    if (!byVisPage) return
    const added: LocalMarkup[] = []
    for (const [visIdx, quads] of byVisPage) {
      const origIdx = visList[visIdx]
      if (origIdx === undefined) continue
      added.push({
        id: `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
        pageIndex: origIdx,
        type,
        color: MARKUP_COLORS[type],
        quads,
      })
    }
    if (added.length === 0) return
    pushUndo()
    setMarkups((prev) => [...prev, ...added])
    window.getSelection()?.removeAllRanges()
  }

  // ── Annotation selection: click selects, deletion is explicit (delete popup / Delete key) ──

  /** Clamp the delete popup anchor into the window, preferring a spot above the click (same rules as the markup bar) */
  const popupPos = (x: number, y: number) => ({
    x: Math.min(Math.max(x, 70), window.innerWidth - 70),
    y: y >= 96 ? y - 48 : Math.min(y + 12, window.innerHeight - 44),
  })

  const deleteSelected = () => {
    const sel = selected
    if (!sel) return
    pushUndo()
    if (sel.kind === 'markup') setMarkups((prev) => prev.filter((m) => m.id !== sel.id))
    else if (sel.kind === 'drawing') setDrawings((prev) => prev.filter((d) => d.id !== sel.id))
    else setStampCfg(null)
    setSelected(null)
    // Transient "deleted · undo" toast so the removal is visible and reversible in place
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current)
    setDeleteToast(true)
    toastTimerRef.current = window.setTimeout(() => setDeleteToast(false), 5000)
  }

  const opFailed = (error: string) => {
    setSaveError(error)
    setSaveState('error')
  }

  /** Pending edits in SavePdfRequest form; shared by in-place Save and Save As */
  const editsPayload = () => ({
    markups: markups.map(({ id: _id, ...rest }) => rest),
    drawings: drawings.map((d) => d.input),
    stamps: stampCfg ? renderStamps(stampCfg, visList) : [],
    formValues: [...formEdits.values()],
    rotations: [...rotations].map(([pageIndex, delta]) => ({ pageIndex, delta })),
    deletedPages: [...deleted],
    ...(order ? { pageOrder: visList } : {}),
    ...(metadata ? { metadata } : {}),
  })

  /** Resolved when the running save() lands; Save As serializes behind it */
  const saveInFlightRef = useRef<Promise<boolean> | null>(null)

  const save = (autosave = false): Promise<boolean> => {
    if (!dirty || saveState === 'saving' || !filePath) return Promise.resolve(!dirty)
    // An explicit save opts this file into autosave
    if (!autosave) savedOnceRef.current = true
    const run = (async (): Promise<boolean> => {
      setSaveState('saving')
      const result = await window.pdfApi.save({ path: filePath, ...editsPayload() })
      if (!result.ok) {
        opFailed(result.error)
        return false
      }
      // Reload: changes are in the file now, canvas renders directly, overlays/pending ops are cleared
      try {
        const el = scrollRef.current
        const scrollTop = el?.scrollTop ?? 0
        await loadDoc(filePath, doc)
        requestAnimationFrame(() => {
          if (scrollRef.current) scrollRef.current.scrollTop = scrollTop
        })
      } catch {
        /* Save already succeeded; a reload failure doesn't block (takes effect on next open) */
      }
      setSaveState('saved')
      setTimeout(() => setSaveState((s) => (s === 'saved' ? 'idle' : s)), 2000)
      return true
    })()
    const tracked = run.finally(() => {
      if (saveInFlightRef.current === tracked) saveInFlightRef.current = null
    })
    saveInFlightRef.current = tracked
    return tracked
  }

  /**
   * Save As: apply pending edits onto the source bytes and write only to targetPath.
   * The original file stays untouched on disk and the edits stay pending in this tab.
   */
  const saveAsTo = async (targetPath: string): Promise<boolean> => {
    if (!filePath) return false
    // A save already in flight (autosave that started before the dialog opened) lands
    // first. If it succeeded, every edit that was pending is now part of the source
    // bytes, so the copy applies nothing on top — deriving this from the save result
    // (instead of re-reading state) avoids racing React's render of the cleared edits.
    const inFlight = saveInFlightRef.current
    const flushed = inFlight ? await inFlight.catch(() => false) : false
    const edits = flushed
      ? { markups: [], drawings: [], formValues: [], stamps: [] }
      : editsPayload()
    setSaveState('saving')
    const result = await window.pdfApi.save({ path: filePath, targetPath, ...edits })
    if (!result.ok) {
      opFailed(result.error)
      return false
    }
    // Back to idle, not 'saved': only the copy was written — this tab's edits are
    // still pending, so a saved-confirmation next to the unsaved badge would lie
    setSaveState('idle')
    return true
  }

  // Autosave pauses while the shell's Save As flow is open: the save dialog blurs the
  // window, and the blur-triggered autosave would write the pending edits into the original
  const saveAsFlowRef = useRef(false)
  useEffect(() => window.pdfApi.onSaveAsFlow((inFlight) => (saveAsFlowRef.current = inFlight)), [])

  // Autosave (same strategy as Docs): every 30s and on window blur, silently persist pending
  // edits via the regular save() path; skipped while a save is in flight or without a file path.
  // Gated on one explicit save first: a PDF opened only to read must never be
  // overwritten because a thumbnail got dragged or a markup tool tapped — Save (⌘S / the
  // toolbar button / File ▸ Save) is what opts this file into unattended writes.
  useAutosave(
    () =>
      savedOnceRef.current &&
      dirty &&
      saveState !== 'saving' &&
      filePath !== '' &&
      !readOnly &&
      !saveAsFlowRef.current,
    () => void save(true),
  )

  // ── Page operations ──

  const rotatePage = (origIdx: number, dir: 90 | -90) => {
    if (readOnly) return
    pushUndo()
    setRotations((prev) => {
      const next = new Map(prev)
      const nv = ((next.get(origIdx) ?? 0) + dir + 360) % 360
      if (nv === 0) next.delete(origIdx)
      else next.set(origIdx, nv)
      return next
    })
  }

  const deletePage = (origIdx: number) => {
    if (pageCount <= 1 || readOnly) return
    pushUndo()
    setDeleted((prev) => new Set(prev).add(origIdx))
    setMarkups((prev) => prev.filter((m) => m.pageIndex !== origIdx))
    setDrawings((prev) => prev.filter((d) => d.input.pageIndex !== origIdx))
  }

  // ── Drawing annotations ──

  const newId = () => `d${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`

  const commitDrawing = (origIdx: number, input: DrawingInput) => {
    pushUndo()
    setDrawings((prev) => [...prev, { id: newId(), input: { ...input, pageIndex: origIdx } }])
  }

  /** Render stamps in current page order; page numbers depend on visList, so both preview and save compute fresh */
  const renderStamps = useCallback(
    (cfg: StampConfig, pages: number[]): StampInput[] =>
      buildStamps(
        pages.map((origIdx, i) => ({
          origIdx,
          pw: sizes[origIdx]!.width,
          ph: sizes[origIdx]!.height,
          displayNo: i + 1,
        })),
        cfg.wm,
        cfg.hf,
      ),
    [sizes],
  )

  const applyStamps = (wm: WatermarkConfig | null, hf: HeaderFooterConfig | null) => {
    setStampDlg(false)
    if (!wm && !hf) return
    pushUndo()
    setStampCfg({ wm, hf })
  }

  /** Stamp preview for visible pages (only pages in rendered rows, so large docs don't render every canvas) */
  const stampPreview = useMemo(() => {
    if (!stampCfg) return new Map<number, StampInput[]>()
    const shown = new Set([...visibleRows].flatMap((r) => rows[r] ?? []))
    const byPage = new Map<number, StampInput[]>()
    for (const s of renderStamps(stampCfg, visList)) {
      if (!shown.has(s.pageIndex)) continue
      const list = byPage.get(s.pageIndex)
      if (list) list.push(s)
      else byPage.set(s.pageIndex, [s])
    }
    return byPage
  }, [stampCfg, visList, rows, visibleRows, renderStamps])

  /** Thumbnail drag-and-drop reorder: move the page at position from to position to */
  const movePage = (from: number, to: number) => {
    if (from === to || readOnly) return
    pushUndo()
    const next = [...visList]
    const [moved] = next.splice(from, 1)
    next.splice(to, 0, moved!)
    // order must cover all original pages (deleted ones included, kept at the tail so they don't affect the result)
    const rest = sizes.map((_, i) => i).filter((i) => !next.includes(i))
    setOrder([...next, ...rest])
  }

  /** Place signature: pad coords → target page PDF coords (width is 1/3 of the page, proportional) */
  const placeSignature = (origIdx: number, at: [number, number]) => {
    const sig = pendingSign
    if (!sig) return
    const geom = pageGeom(origIdx)
    const targetW = geom.pw / 3
    const k = targetW / sig.width
    const paths = sig.paths.map((p) => {
      const out: number[] = []
      for (let i = 0; i < p.length; i += 2) {
        out.push(at[0] + p[i]! * k, at[1] - p[i + 1]! * k)
      }
      return out
    })
    pushUndo()
    setDrawings((prev) => [
      ...prev,
      {
        id: newId(),
        input: { kind: 'ink', pageIndex: origIdx, color: drawColor, width: 1.6, paths },
      },
    ])
    setPendingSign(null)
  }

  /** Export PNG: current page or all visible pages, 150dpi equivalent */
  const exportImages = (allPages: boolean) =>
    flushThen(async () => {
      if (!doc || exporting) return
      setExporting(true)
      try {
        const targets = allPages ? visList : [curOrigIdx].filter((i) => i >= 0)
        const images: string[] = []
        const pageNumbers: number[] = []
        const canvas = document.createElement('canvas')
        for (const origIdx of targets) {
          const page = await doc.getPage(origIdx + 1)
          const viewport = page.getViewport({
            scale: 150 / 72,
            rotation: (page.rotate + rotDelta(origIdx)) % 360,
          })
          canvas.width = Math.floor(viewport.width)
          canvas.height = Math.floor(viewport.height)
          await page.render({ canvas, viewport }).promise
          images.push(canvas.toDataURL('image/png').split(',')[1] ?? '')
          pageNumbers.push(visList.indexOf(origIdx) + 1)
        }
        canvas.width = 0
        canvas.height = 0
        const result = await window.pdfApi.exportImages({
          images,
          pageNumbers,
          baseName: fileName.replace(/\.pdf$/i, ''),
        })
        if (!result.ok) opFailed(result.error)
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    })

  const confirmNote = () => {
    const target = notePrompt
    const text = noteText.trim()
    setNotePrompt(null)
    setNoteText('')
    if (!target || !text) return
    pushUndo()
    setDrawings((prev) => [
      ...prev,
      {
        id: newId(),
        input: {
          kind: 'note',
          pageIndex: target.origIdx,
          color: drawColor,
          at: target.at,
          contents: text,
        },
      },
    ])
  }

  /** Extract/insert work on the file on disk — flush unsaved changes first */
  const flushThen = async (fn: () => Promise<void>) => {
    if (dirty && !(await save())) return
    await fn()
  }

  const extractPage = (origIdx: number) =>
    flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.extractPages({
        path: filePath,
        pages: [origIdx],
        suggestedName: `${base}-p${origIdx + 1}.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })

  const openExtractDlg = () => {
    setExtractInput(String(currentPage))
    setExtractInvalid(false)
    setExtractDlg(true)
  }

  /** Extract dialog confirm: visible page-number ranges → original page indices */
  const confirmExtract = () => {
    const pages = parsePageRanges(extractInput, pageCount)
    if (!pages) {
      setExtractInvalid(true)
      return
    }
    setExtractDlg(false)
    void flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const label = pages.length === 1 ? `p${pages[0]}` : `p${pages[0]}-${pages[pages.length - 1]}`
      const result = await window.pdfApi.extractPages({
        path: filePath,
        pages: pages.map((n) => visList[n - 1]!),
        suggestedName: `${base}-${label}.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })
  }

  const insertPdf = (afterOrigIdx: number) =>
    flushThen(async () => {
      const result = await window.pdfApi.insertPdf({ path: filePath, afterPageIndex: afterOrigIdx })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      if (!('canceled' in result)) await loadDoc(filePath, doc)
    })

  /** Print: save first (markups/forms/page ops all into the file), then reload from the file to render, avoiding a destroyed old doc */
  const printDoc = () =>
    flushThen(async () => {
      if (printing) return
      setPrinting(true)
      try {
        const data = await window.pdfApi.readFile(filePath)
        const pdoc = await getDocument({ data: new Uint8Array(data), ...DOC_OPTS }).promise
        try {
          await printPdf(pdoc)
        } finally {
          void pdoc.destroy()
        }
      } catch (err) {
        opFailed(err instanceof Error ? err.message : String(err))
      } finally {
        setPrinting(false)
      }
    })

  /** Capability surface for AI tools; rebuilt each render (AiPanel mirrors it via refs to get the latest) */
  const aiApi: PdfAiDeps = {
    doc: () => doc,
    fileName: () => fileName,
    pageCount: () => sizes.length,
    currentPage: () => (visList[currentPage - 1] ?? 0) + 1,
    readOnly: () => readOnly,
    outline: () => outline,
    searchIndex: getSearchIndex,
    isDeleted: (i) => deleted.has(i),
    gotoPage: (p) => {
      const visIdx = visList.indexOf(p - 1)
      if (visIdx < 0) return false
      scrollToPage(visIdx + 1)
      return true
    },
    addMarkup: (type, origIdx, rects) => {
      pushUndo()
      const quads = rects.map((r) => [r[0], r[3], r[2], r[3], r[0], r[1], r[2], r[1]])
      setMarkups((prev) => [
        ...prev,
        {
          id: `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
          pageIndex: origIdx,
          type,
          color: MARKUP_COLORS[type],
          quads,
        },
      ])
    },
    formEdits: () => formEdits,
    applyFormEdit: (v) => {
      pushUndo()
      setFormEdits((prev) => new Map(prev).set(v.name, v))
    },
    rotatePage,
    deletePage: (origIdx) => {
      if (pageCount <= 1 || readOnly) return false
      deletePage(origIdx)
      return true
    },
  }

  /** Internal destination of a Link annotation → jump to that page */
  const goToDest = async (dest: unknown) => {
    if (!doc) return
    try {
      const arr = typeof dest === 'string' ? await doc.getDestination(dest) : dest
      if (!Array.isArray(arr)) return
      const ref = arr[0]
      const origIdx =
        typeof ref === 'number'
          ? ref
          : await doc.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])
      const visIdx = visList.indexOf(origIdx)
      if (visIdx >= 0) scrollToPage(visIdx + 1)
    } catch {
      /* Ignore corrupted destinations */
    }
  }

  const curOrigIdx = visList[currentPage - 1] ?? -1

  // Clicking elsewhere closes the thumbnail context menu
  useEffect(() => {
    if (!thumbMenu) return
    const close = (e: PointerEvent) => {
      if (!(e.target as Element | null)?.closest?.('.thumb-menu')) setThumbMenu(null)
    }
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [thumbMenu])

  // Main process picked "Save" in the close prompt → save and report the result
  useEffect(() => {
    return window.pdfApi.onCloseSaveRequest(() => {
      void save().then((ok) => window.pdfApi.sendCloseSaveResult(ok))
    })
  })

  // Shell menu Save As → write pending edits to the picked path only; the original file is never mutated
  useEffect(() => {
    return window.pdfApi.onSaveAsRequest((targetPath) => {
      void saveAsTo(targetPath).then((ok) => window.pdfApi.sendSaveAsResult(ok))
    })
  })

  // Shortcuts: ⌘S/⌘F/⌘P/⌘±/⌘0 + page navigation (only ⌘ combos kept while an input control is focused)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null
      const inEditable =
        !!target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.tagName === 'SELECT' ||
          target.isContentEditable)
      if (e.metaKey || e.ctrlKey) {
        const k = e.key.toLowerCase()
        if (k === 's') {
          e.preventDefault()
          void save()
        } else if (k === 'z') {
          e.preventDefault()
          if (e.shiftKey) redo()
          else undo()
        } else if (k === 'f') {
          e.preventDefault()
          openSearch()
        } else if (k === 'p' && !e.shiftKey) {
          e.preventDefault()
          void printDoc()
        } else if (e.key === '=' || e.key === '+') {
          e.preventDefault()
          zoomIn()
        } else if (e.key === '-') {
          e.preventDefault()
          zoomOut()
        } else if (e.key === '0') {
          e.preventDefault()
          fitModeRef.current = 'width'
          recomputeFit()
        }
        return
      }
      if (e.key === 'Escape') {
        if (pendingSign) setPendingSign(null)
        else if (drawTool) setDrawTool(null)
        else if (selected) setSelected(null)
        else if (searchOpen) closeSearch()
        return
      }
      if (inEditable) return
      if ((e.key === 'Delete' || e.key === 'Backspace') && selected) {
        e.preventDefault()
        deleteSelected()
        return
      }
      const el = scrollRef.current
      if (!el) return
      switch (e.key) {
        case 'PageDown':
        case ' ':
          e.preventDefault()
          el.scrollTop += el.clientHeight - 40
          break
        case 'PageUp':
          e.preventDefault()
          el.scrollTop -= el.clientHeight - 40
          break
        case 'Home':
          e.preventDefault()
          el.scrollTop = 0
          break
        case 'End':
          e.preventDefault()
          el.scrollTop = el.scrollHeight
          break
        case 'ArrowDown':
          e.preventDefault()
          el.scrollTop += 60
          break
        case 'ArrowUp':
          e.preventDefault()
          el.scrollTop -= 60
          break
        case 'ArrowRight':
          e.preventDefault()
          scrollToPage(currentPage + 1)
          break
        case 'ArrowLeft':
          e.preventDefault()
          scrollToPage(currentPage - 1)
          break
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  })

  // Ctrl/⌘ + wheel zoom (native listener: React's wheel is passive and can't preventDefault)
  useEffect(() => {
    const el = scrollRef.current
    if (!el) return
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      e.preventDefault()
      if (e.deltaY < 0) zoomIn()
      else zoomOut()
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  })

  if (status === 'password') {
    return (
      <div className="app">
        <div className="pdf-placeholder">
          <form
            className="pdf-password"
            onSubmit={(e) => {
              e.preventDefault()
              passwordRef.current = pwInput
              setStatus('loading')
              void openPath(filePath)
            }}
          >
            <div className="pdf-password-title">{t('pwTitle')}</div>
            <input
              type="password"
              className="pdf-password-input"
              placeholder={t('pwPlaceholder')}
              value={pwInput}
              autoFocus
              onChange={(e) => setPwInput(e.target.value)}
            />
            {pwWrong && <div className="pdf-password-error">{t('pwWrong')}</div>}
            <button type="submit" className="pdf-password-btn" disabled={!pwInput}>
              {t('pwOpen')}
            </button>
          </form>
        </div>
      </div>
    )
  }

  if (status !== 'ready' || !doc) {
    return (
      <div className="app">
        <div className="pdf-placeholder">
          {status === 'loading' ? t('loading') : status === 'error' ? t('loadError') : t('noFile')}
        </div>
      </div>
    )
  }

  const menuOrig = thumbMenu?.origIdx ?? -1

  return (
    <div className="app">
      <div className="ribbon">
        <div className="ribbon-tabs">
          <button
            className="qa-btn"
            title={`${t('save')} (⌘S)`}
            aria-label={t('save')}
            disabled={!dirty || saveState === 'saving'}
            onClick={() => void save()}
          >
            <IconSave />
          </button>
          <button
            className="qa-btn"
            title={`${t('undo')} (⌘Z)`}
            disabled={undoStack.length === 0}
            onClick={undo}
          >
            <IconUndo />
          </button>
          <button
            className="qa-btn"
            title={`${t('redo')} (⇧⌘Z)`}
            disabled={redoStack.length === 0}
            onClick={redo}
          >
            <IconRedo />
          </button>
          <span className="ribbon-tabs-spacer" />
          <span className="ribbon-file" title={filePath}>
            {fileName}
          </span>
          {readOnly && <span className="tb-readonly">{t('roEncrypted')}</span>}
          {/* Unsaved-changes indicator next to the file name: the file on disk
              is only touched by an explicit save until then */}
          {saveState === 'saving' ? (
            <span className="tb-save-pending">{t('saving')}</span>
          ) : (
            dirty &&
            saveState !== 'error' && <span className="tb-save-pending">{t('unsaved')}</span>
          )}
          {saveState === 'error' && (
            <span className="tb-save-error" title={saveError}>
              {t('saveFailed')}
            </span>
          )}
          {saveState === 'saved' && <span className="tb-save-ok">{t('savedOk')}</span>}
        </div>
        <div className="ribbon-body">
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <button
                className={`rb-big${sidebar === 'thumbs' ? ' active' : ''}`}
                onClick={() => setSidebar((v) => (v === 'thumbs' ? null : 'thumbs'))}
              >
                <span className="rb-big-icon">
                  <IconThumbs />
                </span>
                {t('thumbs')}
              </button>
              <button
                className={`rb-big${sidebar === 'outline' ? ' active' : ''}`}
                disabled={!outline}
                onClick={() => setSidebar((v) => (v === 'outline' ? null : 'outline'))}
              >
                <span className="rb-big-icon">
                  <IconOutline />
                </span>
                {t('outline')}
              </button>
              <button
                className={`rb-big${searchOpen ? ' active' : ''}`}
                title={`${t('search')} (⌘F)`}
                onClick={() => (searchOpen ? closeSearch() : openSearch())}
              >
                <span className="rb-big-icon">
                  <IconSearch />
                </span>
                {t('search')}
              </button>
              <button
                className={`rb-big${spread === 2 ? ' active' : ''}`}
                title={spread === 2 ? t('singlePage') : t('twoPage')}
                onClick={() => setSpread((v) => (v === 1 ? 2 : 1))}
              >
                <span className="rb-big-icon">
                  {spread === 2 ? <IconSinglePage /> : <IconSpread />}
                </span>
                {spread === 2 ? t('singlePage') : t('twoPage')}
              </button>
              <button
                className={`rb-big${nightMode ? ' active' : ''}`}
                title={t('nightMode')}
                onClick={() => setNightMode((v) => !v)}
              >
                <span className="rb-big-icon">
                  <IconNight />
                </span>
                {t('nightMode')}
              </button>
            </div>
          </div>
          <div className="ribbon-sep" />
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <div className="rb-col">
                <div className="rb-row">
                  <input
                    className="tb-page-input"
                    value={pageInput}
                    onChange={(e) => setPageInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && commitPageInput()}
                    onBlur={commitPageInput}
                  />
                  <span className="tb-page-total">{t('pageOf', { total: pageCount })}</span>
                </div>
                <div className="rb-row">
                  <button className="rb-icon" title={t('zoomOut')} onClick={zoomOut}>
                    −
                  </button>
                  <span className="tb-zoom">{Math.round(scale * 100)}%</span>
                  <button className="rb-icon" title={t('zoomIn')} onClick={zoomIn}>
                    +
                  </button>
                </div>
              </div>
              <button
                className="rb-big"
                onClick={() => {
                  fitModeRef.current = 'width'
                  recomputeFit()
                }}
              >
                <span className="rb-big-icon">
                  <IconFitWidth />
                </span>
                {t('fitWidth')}
              </button>
              <button
                className="rb-big"
                onClick={() => {
                  fitModeRef.current = 'page'
                  recomputeFit()
                }}
              >
                <span className="rb-big-icon">
                  <IconFitPage />
                </span>
                {t('fitPage')}
              </button>
            </div>
          </div>
          <div className="ribbon-sep" />
          {/* mousedown preventDefault: the browser clears the text selection the instant the button is pressed, so applyMarkup would lose it */}
          <div className="ribbon-group" onMouseDown={(e) => e.preventDefault()}>
            <div className="ribbon-group-items">
              <button
                className="rb-big"
                disabled={readOnly}
                title={t('highlight')}
                onClick={() => applyMarkup('highlight')}
              >
                <span className="rb-big-icon">
                  <IconHighlight />
                </span>
                {t('highlight')}
              </button>
              <button
                className="rb-big"
                disabled={readOnly}
                title={t('underline')}
                onClick={() => applyMarkup('underline')}
              >
                <span className="rb-big-icon">
                  <IconUnderline />
                </span>
                {t('underline')}
              </button>
              <button
                className="rb-big"
                disabled={readOnly}
                title={t('strikeout')}
                onClick={() => applyMarkup('strikeout')}
              >
                <span className="rb-big-icon">
                  <IconStrike />
                </span>
                {t('strikeout')}
              </button>
            </div>
          </div>
          <div className="ribbon-sep" />
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              {DRAW_TOOLS.map(({ tool, icon: DrawIcon, key }) => (
                <button
                  key={tool}
                  className={`rb-big${drawTool === tool ? ' active' : ''}`}
                  disabled={readOnly}
                  title={t(key)}
                  onClick={() => setDrawTool((v) => (v === tool ? null : tool))}
                >
                  <span className="rb-big-icon">
                    <DrawIcon />
                  </span>
                  {t(key)}
                </button>
              ))}
              <button
                className={`rb-big${pendingSign ? ' active' : ''}`}
                disabled={readOnly}
                title={t('signTitle')}
                onClick={() => (pendingSign ? setPendingSign(null) : setSignDlg(true))}
              >
                <span className="rb-big-icon">
                  <IconSign />
                </span>
                {t('sign')}
              </button>
              <div className="rb-col pdf-color-col">
                <div className="pdf-color-row">
                  {DRAW_COLORS.map((c) => (
                    <button
                      key={c.name}
                      className={`pdf-color-dot${drawColor === c.rgb ? ' active' : ''}`}
                      style={{ background: cssRgb(c.rgb) }}
                      disabled={readOnly}
                      title={t('drawColor')}
                      onClick={() => setDrawColor(c.rgb)}
                    />
                  ))}
                </div>
                <span className="pdf-color-label">{t('drawColor')}</span>
              </div>
            </div>
          </div>
          <div className="ribbon-sep" />
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <button
                className="rb-big"
                disabled={curOrigIdx < 0 || readOnly}
                onClick={() => rotatePage(curOrigIdx, -90)}
              >
                <span className="rb-big-icon">
                  <IconRotateL />
                </span>
                {t('rotateLeft')}
              </button>
              <button
                className="rb-big"
                disabled={curOrigIdx < 0 || readOnly}
                onClick={() => rotatePage(curOrigIdx, 90)}
              >
                <span className="rb-big-icon">
                  <IconRotateR />
                </span>
                {t('rotateRight')}
              </button>
              <button
                className="rb-big"
                disabled={curOrigIdx < 0 || pageCount <= 1 || readOnly}
                onClick={() => deletePage(curOrigIdx)}
              >
                <span className="rb-big-icon">
                  <IconDeletePage />
                </span>
                {t('deletePage')}
              </button>
              <button
                className="rb-big"
                disabled={curOrigIdx < 0 || readOnly}
                onClick={openExtractDlg}
              >
                <span className="rb-big-icon">
                  <IconExtract />
                </span>
                {t('extractPage')}
              </button>
              <button
                className="rb-big"
                disabled={readOnly}
                onClick={() => void insertPdf(curOrigIdx)}
              >
                <span className="rb-big-icon">
                  <IconInsertPdf />
                </span>
                {t('insertPdf')}
              </button>
            </div>
          </div>
          <div className="ribbon-sep" />
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <button
                className="rb-big"
                title={`${t('print')} (⌘P)`}
                disabled={printing}
                onClick={() => void printDoc()}
              >
                <span className="rb-big-icon">
                  <IconPrint />
                </span>
                {printing ? t('printPreparing') : t('print')}
              </button>
              <button
                className="rb-big"
                title={t('exportImagesAll')}
                disabled={exporting}
                onClick={() => void exportImages(true)}
              >
                <span className="rb-big-icon">
                  <IconExportImg />
                </span>
                {exporting ? t('exporting') : t('exportImages')}
              </button>
              <button
                className="rb-big"
                disabled={readOnly}
                title={t('stampTitle')}
                onClick={() => setStampDlg(true)}
              >
                <span className="rb-big-icon">
                  <IconWatermark />
                </span>
                {t('watermark')}
              </button>
              <button className="rb-big" title={t('propsTitle')} onClick={() => setPropsDlg(true)}>
                <span className="rb-big-icon">
                  <IconProps />
                </span>
                {t('props')}
              </button>
            </div>
          </div>
          <div className="ribbon-sep" />
          {/* ---- AI assistant (same far-right Genspark entry as the docs ribbon) ---- */}
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <button
                className={`rb-big rb-ai${aiCollapsed ? '' : ' active'}`}
                title={t('ribbonAiAssistantTip')}
                onClick={() => setAiCollapsed((v) => !v)}
              >
                <span className="rb-big-icon">
                  <GensparkMark size={28} />
                </span>
                <span>{t('ribbonAiAssistant')}</span>
              </button>
            </div>
          </div>
        </div>
      </div>
      <div className="pdf-body">
        {/* dock wrapper animates the width between panel and rail (docs-style 180ms ease);
            the panel stays mounted while collapsed so the chat history survives */}
        <div className={`ai-dock${aiCollapsed ? ' collapsed' : ''}`}>
          {aiCollapsed && (
            <button
              className="ai-rail"
              title={t('aiOpenAssistant')}
              onClick={() => setAiCollapsed(false)}
            >
              <GensparkMark size={22} />
            </button>
          )}
          <AiPanel api={aiApi} onCollapse={() => setAiCollapsed(true)} />
        </div>
        {sidebar === 'outline' && outline && (
          <div className="pdf-thumbs pdf-outline-pane">
            <OutlinePanel outline={outline} onGoToDest={(dest) => void goToDest(dest)} />
          </div>
        )}
        {sidebar === 'thumbs' && (
          <div ref={thumbsRef} className="pdf-thumbs">
            {visList.map((origIdx, v) => {
              const size = dispSize(origIdx)
              return (
                <div
                  key={origIdx}
                  ref={setThumbRef(v)}
                  data-idx={v}
                  className={`pdf-thumb${currentPage === v + 1 ? ' pdf-thumb-active' : ''}${
                    dragOver === v && dragFrom !== null && dragFrom !== v
                      ? ' pdf-thumb-dropbefore'
                      : ''
                  }`}
                  draggable={!readOnly}
                  onDragStart={(e) => {
                    e.dataTransfer.effectAllowed = 'move'
                    setDragFrom(v)
                  }}
                  onDragOver={(e) => {
                    if (dragFrom === null) return
                    e.preventDefault()
                    e.dataTransfer.dropEffect = 'move'
                    setDragOver(v)
                  }}
                  onDragLeave={() => setDragOver((o) => (o === v ? null : o))}
                  onDrop={(e) => {
                    e.preventDefault()
                    if (dragFrom !== null) movePage(dragFrom, v)
                    setDragFrom(null)
                    setDragOver(null)
                  }}
                  onDragEnd={() => {
                    setDragFrom(null)
                    setDragOver(null)
                  }}
                  onClick={() => scrollToPage(v + 1)}
                  onContextMenu={(e) => {
                    e.preventDefault()
                    setThumbMenu({
                      x: Math.min(e.clientX, window.innerWidth - 190),
                      y: Math.min(e.clientY, window.innerHeight - 190),
                      origIdx,
                    })
                  }}
                >
                  <div
                    className="pdf-thumb-box"
                    style={{ height: Math.round((size.height / size.width) * THUMB_WIDTH) }}
                  >
                    <PdfThumb
                      doc={doc}
                      pageNo={origIdx + 1}
                      rotationDelta={rotDelta(origIdx)}
                      visible={visibleThumbs.has(v)}
                    />
                  </div>
                  <span className="pdf-thumb-no">{v + 1}</span>
                </div>
              )
            })}
          </div>
        )}
        <div
          ref={scrollRef}
          className={`pdf-scroll${drawTool ? ' pdf-drawing' : ''}${nightMode ? ' pdf-night' : ''}`}
          onScroll={() => {
            handleScroll()
            setSelPopup(null)
            setSelected(null)
          }}
          onMouseUp={drawTool ? undefined : handleMouseUp}
          onClick={(e) => {
            // Clicking anywhere that isn't an annotation clears the selection
            if (
              !(e.target as Element).closest?.(
                '.pdf-markup, .pdf-draw-shape, .pdf-note-pin, .pdf-stamp-preview',
              )
            )
              setSelected(null)
          }}
        >
          {rows.map((row, r) => (
            <div key={r} ref={setRowRef(r)} data-idx={r} className="pdf-row">
              {row.map((origIdx) => {
                const rowVisible = visibleRows.has(r)
                const size = dispSize(origIdx)
                const geom = pageGeom(origIdx)
                return (
                  <div
                    key={origIdx}
                    className="pdf-page"
                    style={
                      {
                        width: Math.floor(size.width * scale),
                        height: Math.floor(size.height * scale),
                        '--scale-factor': scale,
                      } as CSSProperties
                    }
                  >
                    <PdfPage
                      doc={doc}
                      pageNo={origIdx + 1}
                      scale={scale}
                      rotationDelta={rotDelta(origIdx)}
                      visible={rowVisible}
                    />
                    {pendingSign && (
                      <div
                        className="pdf-sign-drop"
                        title={t('signHint')}
                        onClick={(e) => {
                          const box = e.currentTarget.getBoundingClientRect()
                          placeSignature(
                            origIdx,
                            viewToPdf(
                              geom,
                              (e.clientX - box.left) / scale,
                              (e.clientY - box.top) / scale,
                            ),
                          )
                        }}
                      />
                    )}
                    {rowVisible && (
                      <>
                        {/* Preview of unsaved stamps; clicking selects the whole watermark/header-footer set */}
                        {(stampPreview.get(origIdx) ?? []).map((s, si) => (
                          <img
                            key={si}
                            className={`pdf-stamp-preview${selected?.kind === 'stamp' ? ' pdf-stamp-selected' : ''}`}
                            src={`data:image/png;base64,${s.image}`}
                            alt=""
                            title={t('removeStamp')}
                            style={{
                              ...pdfRectToCss(geom, s.rect, scale),
                              opacity: s.opacity ?? 1,
                            }}
                            onClick={(e) =>
                              setSelected({ kind: 'stamp', ...popupPos(e.clientX, e.clientY) })
                            }
                          />
                        ))}
                        {searchOpen && (
                          <div className="pdf-search-layer">
                            {activeMatches.flatMap((m, mi) =>
                              m.pageIndex === origIdx
                                ? m.rects.map((r, ri) => (
                                    <div
                                      key={`${mi}-${ri}`}
                                      className={`pdf-search-hit${mi === searchCurClamped ? ' pdf-search-hit-cur' : ''}`}
                                      style={pdfRectToCss(geom, r, scale)}
                                    />
                                  ))
                                : [],
                            )}
                          </div>
                        )}
                        <MarkupOverlay
                          markups={markups.filter((m) => m.pageIndex === origIdx)}
                          geom={geom}
                          scale={scale}
                          selectedId={selected?.kind === 'markup' ? selected.id : null}
                          selectTitle={t('removeMarkup')}
                          onSelect={(id, x, y) =>
                            setSelected({ kind: 'markup', id, ...popupPos(x, y) })
                          }
                        />
                        <DrawLayer
                          geom={geom}
                          scale={scale}
                          pageWidth={size.width}
                          pageHeight={size.height}
                          drawings={drawings.filter((d) => d.input.pageIndex === origIdx)}
                          tool={readOnly ? null : drawTool}
                          color={drawColor}
                          strokeWidth={STROKE_WIDTH}
                          selectedId={selected?.kind === 'drawing' ? selected.id : null}
                          selectTitle={t('removeMarkup')}
                          onCommit={(input) => commitDrawing(origIdx, input)}
                          onNoteAt={(at) => {
                            setNoteText('')
                            setNotePrompt({ origIdx, at })
                          }}
                          onSelect={(id, x, y) =>
                            setSelected({ kind: 'drawing', id, ...popupPos(x, y) })
                          }
                        />
                        <LinkLayer
                          doc={doc}
                          pageNo={origIdx + 1}
                          geom={geom}
                          scale={scale}
                          onGoToDest={(dest) => void goToDest(dest)}
                        />
                        <FormLayer
                          doc={doc}
                          pageNo={origIdx + 1}
                          geom={geom}
                          scale={scale}
                          readOnly={readOnly}
                          edits={formEdits}
                          onEdit={(v2) => {
                            pushUndo(`form:${v2.name}`)
                            setFormEdits((prev) => new Map(prev).set(v2.name, v2))
                          }}
                        />
                      </>
                    )}
                  </div>
                )
              })}
            </div>
          ))}
        </div>
        {searchOpen && (
          <div className="pdf-search-bar">
            <input
              ref={searchInputRef}
              className="pdf-search-input"
              placeholder={t('search')}
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') searchStep(e.shiftKey ? -1 : 1)
                else if (e.key === 'Escape') closeSearch()
              }}
            />
            <span className="pdf-search-count">
              {searchQuery.trim()
                ? activeMatches.length > 0
                  ? t('searchCount', { current: searchCurClamped + 1, total: activeMatches.length })
                  : t('searchNoResults')
                : ''}
            </span>
            <button
              className="rb-icon"
              title={t('searchPrev')}
              disabled={activeMatches.length === 0}
              onClick={() => searchStep(-1)}
            >
              ‹
            </button>
            <button
              className="rb-icon"
              title={t('searchNext')}
              disabled={activeMatches.length === 0}
              onClick={() => searchStep(1)}
            >
              ›
            </button>
            <button className="rb-icon" onClick={closeSearch}>
              ×
            </button>
          </div>
        )}
        {selPopup && (
          <div
            className="pdf-sel-popup"
            style={{ left: selPopup.x, top: selPopup.y }}
            onMouseDown={(e) => e.preventDefault()}
          >
            <button type="button" title={t('highlight')} onClick={() => applyMarkup('highlight')}>
              <span className="sel-swatch sel-swatch-hl" />
            </button>
            <button type="button" title={t('underline')} onClick={() => applyMarkup('underline')}>
              <span className="sel-swatch sel-swatch-ul">U</span>
            </button>
            <button type="button" title={t('strikeout')} onClick={() => applyMarkup('strikeout')}>
              <span className="sel-swatch sel-swatch-st">S</span>
            </button>
          </div>
        )}
        {selected && (
          <div
            className="pdf-del-popup"
            style={{ left: selected.x, top: selected.y }}
            onMouseDown={(e) => e.preventDefault()}
          >
            <button type="button" onClick={deleteSelected}>
              {t('deleteAnnotation')}
            </button>
          </div>
        )}
        {deleteToast && (
          <div className="pdf-toast">
            <span>{t('annotationDeleted')}</span>
            <button
              type="button"
              onClick={() => {
                setDeleteToast(false)
                undo()
              }}
            >
              {t('undo')}
            </button>
          </div>
        )}
        {thumbMenu && (
          <div className="thumb-menu file-menu" style={{ left: thumbMenu.x, top: thumbMenu.y }}>
            <button
              onClick={() => {
                rotatePage(menuOrig, -90)
                setThumbMenu(null)
              }}
            >
              {t('rotateLeft')}
            </button>
            <button
              onClick={() => {
                rotatePage(menuOrig, 90)
                setThumbMenu(null)
              }}
            >
              {t('rotateRight')}
            </button>
            <button
              disabled={pageCount <= 1}
              onClick={() => {
                deletePage(menuOrig)
                setThumbMenu(null)
              }}
            >
              {t('deletePage')}
            </button>
            <button
              onClick={() => {
                setThumbMenu(null)
                void extractPage(menuOrig)
              }}
            >
              {t('extractPage')}
            </button>
            <button
              onClick={() => {
                setThumbMenu(null)
                void insertPdf(menuOrig)
              }}
            >
              {t('insertPdf')}
            </button>
          </div>
        )}
        {stampDlg && (
          <StampDialog t={t} onCancel={() => setStampDlg(false)} onApply={applyStamps} />
        )}
        {propsDlg && (
          <PropertiesDialog
            doc={doc}
            fileName={fileName}
            fileSize={fileSize}
            pageCount={pageCount}
            pending={metadata}
            readOnly={readOnly}
            t={t}
            onCancel={() => setPropsDlg(false)}
            onApply={(meta) => {
              setPropsDlg(false)
              pushUndo()
              setMetadata(meta)
            }}
          />
        )}
        {signDlg && (
          <SignatureDialog
            color={drawColor}
            t={t}
            onCancel={() => setSignDlg(false)}
            onConfirm={(sig) => {
              setSignDlg(false)
              setPendingSign(sig)
            }}
          />
        )}
        {notePrompt && (
          <div className="pdf-modal-mask" onClick={() => setNotePrompt(null)}>
            <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
              <div className="pdf-modal-title">{t('noteTitle')}</div>
              <textarea
                className="pdf-modal-textarea"
                value={noteText}
                placeholder={t('notePlaceholder')}
                autoFocus
                onChange={(e) => setNoteText(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) confirmNote()
                  else if (e.key === 'Escape') setNotePrompt(null)
                }}
              />
              <div className="pdf-modal-actions">
                <button className="pdf-modal-btn" onClick={() => setNotePrompt(null)}>
                  {t('cancel')}
                </button>
                <button
                  className="pdf-modal-btn primary"
                  disabled={!noteText.trim()}
                  onClick={confirmNote}
                >
                  {t('ok')}
                </button>
              </div>
            </div>
          </div>
        )}
        {extractDlg && (
          <div className="pdf-modal-mask" onClick={() => setExtractDlg(false)}>
            <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
              <div className="pdf-modal-title">{t('extractRangeTitle')}</div>
              <input
                className={`pdf-modal-input${extractInvalid ? ' invalid' : ''}`}
                value={extractInput}
                placeholder={t('extractRangeHint', { total: pageCount })}
                autoFocus
                onChange={(e) => {
                  setExtractInput(e.target.value)
                  setExtractInvalid(false)
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') confirmExtract()
                  else if (e.key === 'Escape') setExtractDlg(false)
                }}
              />
              <div className="pdf-modal-actions">
                <button className="pdf-modal-btn" onClick={() => setExtractDlg(false)}>
                  {t('cancel')}
                </button>
                <button className="pdf-modal-btn primary" onClick={confirmExtract}>
                  {t('ok')}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
