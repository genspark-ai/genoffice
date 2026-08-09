import { Fragment, useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type {
  CSSProperties,
  MouseEvent as ReactMouseEvent,
  ReactElement,
  ReactNode,
  RefObject,
} from 'react'
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
import { groupLineSpans } from './text-line'
import { DRAW_COLORS, DrawLayer, cssRgb } from './DrawLayer'
import type { DrawTool, LocalDrawing } from './DrawLayer'
import { FormLayer } from './FormLayer'
import { ImageEditLayer, imageRectKey } from './ImageEditLayer'
import type { LocalImageEdit } from './ImageEditLayer'
import { navAction } from './keyNav'
import { rowOfVisIdx, spreadRows, stepPage } from './spread'
import { LinkLayer } from './LinkLayer'
import { OutlinePanel } from './OutlinePanel'
import type { OutlineNode } from './OutlinePanel'
import { printPdf } from './print'
import { PropertiesDialog } from './PropertiesDialog'
import { SignatureDialog, fileToCanvas } from './SignatureDialog'
import type { SignatureData } from './SignatureDialog'
import { StampDialog } from './StampDialog'
import { buildStamps } from './stamps'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import { buildSearchIndex, searchInIndex } from './search'
import type { SearchIndex, SearchMatch } from './search'
import { groupPageBlocks, type TextBlock } from './text-block'
import { joinBlockLines, measurePt, wrapText } from './text-wrap'
import {
  colorRunsEqual,
  colorSegments,
  colorsToRuns,
  mapCharColors,
  runsToColors,
  spliceCharColors,
} from './color-runs'
import { useI18n } from './i18n/locale'
import { useAutosave } from './useAutosave'
import { EDIT_FONTS } from '../shared/ipc'
import type {
  DrawingInput,
  FormValueInput,
  ImageEditFailure,
  ImageLayer,
  MarkupType,
  MetadataInput,
  PageImageRef,
  StampInput,
  TextEditFailure,
  TextEditInput,
} from '../shared/ipc'

const EDIT_FONT_BY_ID = new Map<string, (typeof EDIT_FONTS)[number]>(
  EDIT_FONTS.map((f) => [f.id, f]),
)

GlobalWorkerOptions.workerSrc = workerUrl

// cmaps/standard fonts/wasm are statically copied by the build into pdfjs/ of the renderer output (same path on the dev server)
const ASSET_BASE = new URL('pdfjs/', document.baseURI).href

let measureCtx: CanvasRenderingContext2D | null = null
/** Width of text in the given CSS font (shared hidden canvas) */
function measureTextWidth(text: string, font: string): number {
  measureCtx ??= document.createElement('canvas').getContext('2d')
  if (!measureCtx) return 0
  measureCtx.font = font
  return measureCtx.measureText(text).width
}

const ZOOM_STEPS = [0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]
const MIN_SCALE = ZOOM_STEPS[0]
const MAX_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]
const PAGE_GAP = 16
const SCROLL_PAD = 24
// ── Sidebar (thumbnails / outline) width: drag the divider to resize; persisted ──
const SIDEBAR_W_KEY = 'genoffice-pdf-sidebar-width'
const SIDEBAR_W_DEFAULT = 150
const SIDEBAR_W_MIN = 120
/** pane padding (10px × 2) + thumb box borders (2px × 2) */
const SIDEBAR_CHROME = 24

const clampSidebarW = (w: number): number =>
  Math.min(Math.max(w, SIDEBAR_W_MIN), Math.min(320, Math.round(window.innerWidth * 0.4)))

const loadSidebarW = (): number => {
  const saved = Number(localStorage.getItem(SIDEBAR_W_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampSidebarW(saved) : SIDEBAR_W_DEFAULT
}

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

/** Thumbnail: rendered once per (doc, rotation, raster width) when visible and cached.
 *  rasterW only changes when a sidebar drag ends, so a resize re-rasters each
 *  visible thumb once; while dragging the canvas just CSS-stretches. */
function PdfThumb({
  doc,
  pageNo,
  rotationDelta,
  visible,
  rasterW,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  rotationDelta: number
  visible: boolean
  rasterW: number
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const renderedKeyRef = useRef<string | null>(null)
  useEffect(() => {
    const canvas = canvasRef.current
    const key = `${rotationDelta}:${rasterW}`
    if (!visible || !canvas || renderedKeyRef.current === key) return
    renderedKeyRef.current = key
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const rotation = (page.rotate + rotationDelta) % 360
      const scale = rasterW / page.getViewport({ scale: 1, rotation }).width
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
  }, [doc, pageNo, rotationDelta, visible, rasterW])
  // Reset the cache key when the doc changes (save reload); re-render next time it's visible
  useEffect(() => {
    renderedKeyRef.current = null
  }, [doc])
  return <canvas ref={canvasRef} style={{ width: '100%' }} />
}

// ── ribbon icons (aligned with slides' rb-big visual language) ──

/** Constant painted stroke instead of proportional scaling — same rule as the
 *  slides icons: ~1.5px lines at 20px+, ~1.25px on 13-19px glyphs, ~1.1px below.
 *  stroke-width is in 24-canvas units: units = painted-px × 24 / rendered-px. */
function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 24) / size
}

function Icon({ size = 28, children }: { size?: number; children: ReactNode }): ReactElement {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
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
    <rect x="4.5" y="5" width="6" height="6.5" rx="1" />
    <rect x="4.5" y="14" width="6" height="5" rx="1" />
    <path d="M14 6 L19.5 6 M14 10 L19.5 10 M14 15 L19.5 15 M14 18 L19.5 18" />
  </Icon>
)
const IconHighlight = () => (
  <Icon>
    <path d="M6.04 15.09 L13.54 7.59 L16.63 10.68 L9.13 18.18 L6.04 18.18 L6.04 15.09 Z" />
    <path d="M13.54 6.71 L15.75 4.5 L18.84 7.59 L16.63 9.79" />
    <path d="M5.16 19.5 L17.51 19.5" strokeWidth={2.2} />
  </Icon>
)
const IconUnderline = () => (
  <Icon>
    <path d="M7.56 4.5 L7.56 11.17 A 4.44 4.44 0 0 0 16.44 11.17 L16.44 4.5" />
    <path d="M5.89 19.5 L18.11 19.5" strokeWidth={1.85} />
  </Icon>
)
const IconStrike = () => (
  <Icon>
    <path d="M16.29 7.21 C15.72 5.52 14.03 4.5 12 4.5 C9.52 4.5 7.71 5.85 7.71 7.88 C7.71 9.46 8.73 10.36 10.65 10.93" />
    <path d="M8.62 16.91 C9.18 18.48 10.87 19.5 13.02 19.5 C15.5 19.5 17.41 18.26 17.41 16.12 C17.41 15.44 17.3 14.88 16.96 14.42" />
    <path d="M4.67 12.28 L19.33 12.28" strokeWidth={1.65} />
  </Icon>
)
const IconEditText = () => (
  <Icon>
    <path d="M5 5.5 L14.5 5.5 M9.75 5.5 L9.75 15.5 M7.5 15.5 L12 15.5" />
    <path d="M16.7 10.3 L19.7 13.3 L13.4 19.6 L10.4 20.1 L10.9 17.1 Z" />
  </Icon>
)
const IconInk = () => (
  <Icon>
    <path d="M16.15 4.85 L19.15 7.85 L8.9 18.1 L4.9 19.1 L5.9 15.1 Z" />
    <path d="M14.15 6.85 L17.15 9.85" />
  </Icon>
)
const IconRect = () => (
  <Icon>
    <rect x="4.5" y="6.64" width="15" height="10.71" rx="1.07" />
  </Icon>
)
const IconEllipse = () => (
  <Icon>
    <ellipse cx="12" cy="12" rx="7.5" ry="5.57" />
  </Icon>
)
const IconArrow = () => (
  <Icon>
    <path d="M4.5 19.2 L19.5 4.8" />
    <path d="M12.9 4.8 L19.5 4.8 L19.5 11.4" />
  </Icon>
)
const IconNote = () => (
  <Icon>
    <path d="M4.5 5.57 L19.5 5.57 L19.5 15.21 L10.93 15.21 L6.64 18.43 L6.64 15.21 L4.5 15.21 Z" />
    <path d="M8.25 9.32 L15.75 9.32 M8.25 12 L13.07 12" />
  </Icon>
)
const IconSign = () => (
  <Icon>
    <path d="M5.5 15.1 C7.8 12.3 9.5 9 9.2 7 C9 5.7 7.9 5.9 7.6 7.4 C7.2 9.6 8.6 13.4 10.5 14.9 C12 16.1 13.9 15.3 14.7 13.8 C15.1 13 15.9 13 16.3 13.8 C16.7 14.7 17.7 15 18.5 14.4" />
    <path d="M4.75 18.6 L19.25 18.6" />
  </Icon>
)
const IconExportImg = () => (
  <Icon>
    <rect x="4.5" y="6.75" width="15" height="10.5" rx="1" />
    <circle cx="9" cy="10.55" r="1.2" />
    <path d="M4.8 14.95 L9 11.75 L12.4 14.35 L15 12.35 L19.2 15.75" />
  </Icon>
)
const IconInsertImage = () => (
  <Icon>
    <rect x="4.5" y="6" width="11.5" height="9.5" rx="1" />
    <circle cx="8" cy="9.2" r="1.1" />
    <path d="M4.8 13.6 L8 11.2 L11 13.4 L13.2 11.8 L15.8 13.9" />
    <path d="M18.6 13.4 V19 M15.8 16.2 H21.4" />
  </Icon>
)
const IconEditImage = () => (
  <Icon>
    <rect x="4.5" y="6" width="12.5" height="10" rx="1" />
    <circle cx="8.3" cy="9.4" r="1.1" />
    <path d="M4.8 14 L8.3 11.4 L11.5 13.7 L13.8 12" />
    <path d="M14.2 18.9 L19.7 13.4 A1.06 1.06 0 0 0 18.2 11.9 L12.7 17.4 L12.2 19.4 Z" />
  </Icon>
)
const IconNight = () => (
  <Icon>
    <path d="M19.5 13.48 A 7.58 7.58 0 0 1 10.52 4.5 A 7.58 7.58 0 1 0 19.5 13.48 Z" />
  </Icon>
)
const IconSpread = () => (
  <Icon>
    <rect x="4.5" y="6" width="6.5" height="12" rx="1" />
    <rect x="13" y="6" width="6.5" height="12" rx="1" />
  </Icon>
)
const IconSinglePage = () => (
  <Icon>
    <rect x="6.81" y="4.5" width="10.38" height="15" rx="1.15" />
  </Icon>
)
const IconWatermark = () => (
  <Icon>
    <rect x="4.5" y="5.04" width="15" height="13.93" rx="1.07" />
    <path d="M7.71 15.75 L15.75 7.71" />
    <path d="M7.71 11.46 L11.46 7.71 M12.54 15.75 L16.29 12" />
  </Icon>
)
const IconProps = () => (
  <Icon>
    <circle cx="12" cy="12" r="7.5" />
    <path d="M12 10.96 L12 15.65" />
    <circle cx="12" cy="8.46" r="0.94" fill="currentColor" stroke="none" />
  </Icon>
)
const IconRotateL = () => (
  <Icon>
    <path d="M8.28 10.3 L4.53 10.3 L4.53 6.55" />
    <path d="M4.75 9.98 A 7.5 7.5 0 1 1 4.53 12.98" />
  </Icon>
)
const IconRotateR = () => (
  <Icon>
    <path d="M15.72 10.3 L19.47 10.3 L19.47 6.55" />
    <path d="M19.25 9.98 A 7.5 7.5 0 1 0 19.47 12.98" />
  </Icon>
)
const IconDeletePage = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M9.7 11.75 L14.2 16.25 M14.2 11.75 L9.7 16.25" />
  </Icon>
)
const IconExtract = () => (
  <Icon>
    <path d="M7.2 4.5 H13.2 L16.7 8 V11.5" />
    <path d="M6.2 5.5 V18.5 A1 1 0 0 0 7.2 19.5 H11.2" />
    <path d="M14.95 13.5 V19 M12.2 16.5 L14.95 19.25 L17.7 16.5" />
  </Icon>
)
const IconInsertPdf = () => (
  <Icon>
    <path d="M7.7 4.5 H13.7 L17.2 8 V18.5 A1 1 0 0 1 16.2 19.5 H7.7 A1 1 0 0 1 6.7 18.5 V5.5 A1 1 0 0 1 7.7 4.5 Z" />
    <path d="M13.7 4.5 V8 H17.2" />
    <path d="M11.95 11 V17 M8.95 14 H14.95" />
  </Icon>
)
const IconFitWidth = () => (
  <Icon>
    <path d="M4.5 5.57 L4.5 18.43 M19.5 5.57 L19.5 18.43" />
    <path d="M7.71 12 L16.29 12 M10.07 9.64 L7.71 12 L10.07 14.36 M13.93 9.64 L16.29 12 L13.93 14.36" />
  </Icon>
)
const IconFitPage = () => (
  <Icon>
    <rect x="7" y="4.5" width="10" height="15" rx="1" />
    <path d="M12 8 L12 16 M9.8 10.2 L12 8 L14.2 10.2 M9.8 13.8 L12 16 L14.2 13.8" />
  </Icon>
)
const IconOutline = () => (
  <Icon>
    <path d="M4.84 4.78 L19.5 4.78 M8.22 9.29 L19.5 9.29 M8.22 13.8 L19.5 13.8 M11.61 18.32 L19.5 18.32" />
    <circle cx="5.4" cy="9.29" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="5.4" cy="13.8" r="0.9" fill="currentColor" stroke="none" />
    <circle cx="8.79" cy="18.32" r="0.9" fill="currentColor" stroke="none" />
  </Icon>
)
const IconDrawColor = () => (
  <Icon>
    <path d="M12 4.5 C14.2 7.3 17.25 9.2 17.25 12.4 C17.25 15.4 14.9 17.5 12 17.5 C9.1 17.5 6.75 15.4 6.75 12.4 C6.75 9.2 9.8 7.3 12 4.5 Z" />
  </Icon>
)
/* dropdown chevron, same glyph as slides' RbCaret */
const RbCaret = () => (
  <svg className="rb-caret" width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden>
    <path
      d="M5.5 9.25 12 15.75l6.5-6.5"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
)
const IconSearch = () => (
  <Icon>
    <circle cx="10.61" cy="10.61" r="6.11" />
    <path d="M15.28 15.28 L19.5 19.5" />
  </Icon>
)
const IconPrint = () => (
  <Icon>
    <path d="M7.71 8.79 L7.71 4.5 L16.29 4.5 L16.29 8.79" />
    <rect x="5.04" y="8.79" width="13.93" height="6.96" rx="1.07" />
    <path d="M7.71 13.07 L16.29 13.07 L16.29 19.5 L7.71 19.5 Z" />
  </Icon>
)
/** Design-supplied glyphs on the 1:16 stroke:canvas ratio (24-canvas / 1.5 stroke),
 *  geometry shared with docs/slides/sheets. */
const IconRatio = ({ size = 16, children }: { size?: number; children: ReactNode }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    {children}
  </svg>
)
const IconUndo = () => (
  <IconRatio>
    <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
    <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
  </IconRatio>
)
const IconRedo = () => (
  <IconRatio>
    <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
    <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
  </IconRatio>
)
const IconSave = () => (
  <IconRatio>
    <path d="M3 4.5C3 3.67158 3.67158 3 4.5 3H17.1407L21 6.60325V19.5C21 20.3285 20.3285 21 19.5 21H4.5C3.67158 21 3 20.3285 3 19.5V4.5Z" />
    <path d="M12.0042 3L12 6.6923C12 6.86225 11.7761 7 11.5 7H7.5C7.22385 7 7 6.86225 7 6.6923V3H12.0042Z" />
    <path d="M7 13H17" />
    <path d="M7 17H12.0042" />
  </IconRatio>
)

// ── selection-popup icons (14px; bring-forward / send-backward / trash) ──

const IconLayerUp = () => (
  <Icon size={14}>
    <rect x="9.5" y="4.5" width="10" height="10" rx="1" />
    <path d="M14.5 19.5 H5.5 A1 1 0 0 1 4.5 18.5 V9.5" />
  </Icon>
)
const IconLayerDown = () => (
  <Icon size={14}>
    <rect x="4.5" y="9.5" width="10" height="10" rx="1" />
    <path d="M9.5 4.5 H18.5 A1 1 0 0 1 19.5 5.5 V14.5" />
  </Icon>
)
const IconTrash = () => (
  <Icon size={14}>
    <path d="M4.5 6.5 H19.5" />
    <path d="M9 6.5 V5 A1 1 0 0 1 10 4 H14 A1 1 0 0 1 15 5 V6.5" />
    <path d="M6.5 6.5 L7.3 18.6 A1.4 1.4 0 0 0 8.7 19.9 H15.3 A1.4 1.4 0 0 0 16.7 18.6 L17.5 6.5" />
    <path d="M10.2 10 V16 M13.8 10 V16" />
  </Icon>
)
const IconRotateCw = () => (
  <Icon size={14}>
    <path d="M18.5 8.5 A7.5 7.5 0 1 0 19.5 12" />
    <path d="M19 4 V8.5 H14.5" />
  </Icon>
)
const IconRotateCcw = () => (
  <Icon size={14}>
    <path d="M5.5 8.5 A7.5 7.5 0 1 1 4.5 12" />
    <path d="M5 4 V8.5 H9.5" />
  </Icon>
)
const IconSwapImage = () => (
  <Icon size={14}>
    <rect x="4" y="6.5" width="11" height="11" rx="1" />
    <circle cx="7.5" cy="10" r="1.2" />
    <path d="M4.5 16 L8.5 12.5 L11 15 L12.5 13.5 L15 16" />
    <path d="M17 4.5 H19 A1 1 0 0 1 20 5.5 V13" />
    <path d="M18.2 11.2 L20 13 L21.8 11.2" />
  </Icon>
)

const rgbToHex = (c: readonly [number, number, number]): string =>
  `#${c
    .map((v) =>
      Math.round(v * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`
const hexToRgb = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16) / 255,
  parseInt(hex.slice(3, 5), 16) / 255,
  parseInt(hex.slice(5, 7), 16) / 255,
]

/** Text-edit colors travel as 0-255 RGB (PDFium fill color), unlike markups' 0-1 floats */
const hexTo255 = (hex: string): [number, number, number] => [
  parseInt(hex.slice(1, 3), 16),
  parseInt(hex.slice(3, 5), 16),
  parseInt(hex.slice(5, 7), 16),
]
const rgb255ToHex = (c: readonly [number, number, number]): string =>
  `#${c.map((v) => v.toString(16).padStart(2, '0')).join('')}`

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

// ── ribbon tabs (docs-style tab strip over a fixed 80px band) ──
const RIBBON_TABS = [
  { id: 'home', labelKey: 'ribbonTabHome' },
  { id: 'annotate', labelKey: 'ribbonTabAnnotate' },
  { id: 'edit', labelKey: 'ribbonTabEdit' },
  { id: 'page', labelKey: 'ribbonTabPage' },
  { id: 'view', labelKey: 'ribbonTabView' },
] as const
type RibbonTab = (typeof RIBBON_TABS)[number]['id']

/** One-click AI feature glyphs (same 24-canvas/1.5-stroke artwork as the docs ribbon) */
const IconAiSummarize = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M13.875 21H12H6.5C5.39543 21 4.5 20.1046 4.5 19V5C4.5 3.89543 5.39543 3 6.5 3H17.5C18.6046 3 19.5 3.89543 19.5 5V9V12V13" />
    <path d="M8.00001 7H16" />
    <path d="M8.00007 10.2032H14.0001" />
    <path d="M8.00007 13.4062H12.0001" />
    <path
      d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
      strokeLinejoin="round"
    />
  </svg>
)
const IconAiKeyPoints = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.5"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden
  >
    <path d="M4 5H20" />
    <path d="M4 9H16" />
    <path d="M4 13H11" />
    <path d="M4 17H10" />
    <path
      d="M17 14L17.2579 14.697C17.5961 15.611 17.7652 16.068 18.0986 16.4014C18.432 16.7348 18.889 16.9039 19.803 17.2421L20.5 17.5L19.803 17.7579C18.889 18.0961 18.432 18.2652 18.0986 18.5986C17.7652 18.932 17.5961 19.389 17.2579 20.303L17 21L16.7421 20.303C16.4039 19.389 16.2348 18.932 15.9014 18.5986C15.568 18.2652 15.111 18.0961 14.197 17.7579L13.5 17.5L14.197 17.2421C15.111 16.9039 15.568 16.7348 15.9014 16.4014C16.2348 16.068 16.4039 15.611 16.7421 14.697L17 14Z"
      strokeLinejoin="round"
    />
  </svg>
)

/** Drawing stroke width (PDF pt); thin lines stay crisp under zoom */
const STROKE_WIDTH = 2

/** Full snapshot of unsaved edits (for undo/redo; data is small, whole-copy replace is safest) */
/** Watermark/header-footer are kept as config and rendered in final page order only at save time, so page numbers survive reorders/deletions */
interface StampConfig {
  wm: WatermarkConfig | null
  hf: HeaderFooterConfig | null
}

interface LocalTextEdit {
  id: string
  input: TextEditInput
  /** Matched run's ink bounds from validation (PDF user space). The edit rect is a pdf.js
      layout box; glyph ink can poke out of it, so the preview covers this instead */
  cover?: [number, number, number, number]
}

/** Area a pending edit must blank: the edit rect grown to the validated ink bounds */
const unionCover = (
  rect: readonly [number, number, number, number],
  cover: readonly [number, number, number, number] | undefined,
): [number, number, number, number] =>
  cover
    ? [
        Math.min(rect[0], cover[0]),
        Math.min(rect[1], cover[1]),
        Math.max(rect[2], cover[2]),
        Math.max(rect[3], cover[3]),
      ]
    : [rect[0], rect[1], rect[2], rect[3]]

/** Expand a CSS box by p px on every side (antialiasing bleeds past exact ink bounds) */
const inflateCss = (
  b: { left: number; top: number; width: number; height: number },
  p: number,
) => ({
  left: b.left - p,
  top: b.top - p,
  width: b.width + 2 * p,
  height: b.height + 2 * p,
})

/** Editor state for the floating text-edit box; editId set when re-opening a pending edit */
interface TextDraft {
  origIdx: number
  rect: [number, number, number, number]
  oldText: string
  fontSize: number
  value: string
  /** Style overrides; undefined = keep the run's original size/color */
  size?: number
  /** CSS hex like '#d32f2f' */
  color?: string
  /** Selection-level colors, one hex per code unit of value; '' = base (color ?? original).
      undefined/all-'' = uniform draft (the pre-existing single-color behavior). */
  charColors?: string[]
  /** EDIT_FONTS id; undefined = automatic rebuild font */
  font?: string
  /** Style toggles; true = on, undefined = off (resolved via font variants at save) */
  bold?: true
  italic?: true
  editId?: string
  /** Ink bounds of the run being edited (async, from a dry-run validate) */
  cover?: [number, number, number, number]
  /** Present for paragraph (block) edits: the geometry the commit reflows into.
      lineHeight is the block's original leading at the original font size. */
  block?: {
    leftPt: number
    firstBaseline: number
    widthPt: number
    lineHeight: number
    align: 'left' | 'center' | 'right'
  }
}

interface EditSnapshot {
  markups: LocalMarkup[]
  drawings: LocalDrawing[]
  textEdits: LocalTextEdit[]
  imageEdits: LocalImageEdit[]
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  deleted: Set<number>
  order: number[] | null
  metadata: MetadataInput | null
}

/** What a running save wrote, captured when the save starts. The post-save reload
    subtracts exactly this instead of wiping all edit state, so anything the user did
    while the write was in flight stays pending on the reloaded document. */
interface SavedSnapshot {
  markupIds: Set<string>
  drawingIds: Set<string>
  textEditIds: Set<string>
  imageEditIds: Set<string>
  stampCfg: StampConfig | null
  formEdits: Map<string, FormValueInput>
  rotations: Map<number, number>
  metadata: MetadataInput | null
  /** Old original page index → its index in the saved file (saved deletions/reorder applied) */
  pageMap: Map<number, number>
}

/** Selected annotation with the anchor of its floating delete popup; a stamp click selects the whole watermark/header-footer set */
type AnnotSelection =
  | { kind: 'markup' | 'drawing' | 'textEdit' | 'imageEdit'; id: string; x: number; y: number }
  | { kind: 'pageImage'; ref: PageImageRef; x: number; y: number }
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

/** Scale factor applied when placing a signature: 1/3 of the displayed page width,
    capped at 1/6 of its height so tall images stay signature-sized */
const signPlaceK = (sig: SignatureData, dispW: number, dispH: number): number =>
  Math.min(dispW / 3 / sig.width, dispH / 6 / sig.height)

/** Inserted images land at up to half the page, never above natural size
    (0.75 ≈ px→pt, so a screen-resolution image keeps its printed size) */
const imagePlaceK = (sig: SignatureData, dispW: number, dispH: number): number =>
  Math.min(dispW / 2 / sig.width, dispH / 2 / sig.height, 0.75)

/** Click-to-place overlay: a translucent ghost of the pending signature follows the
    cursor at its actual landing size, and clicking drops it centered on that point */
function SignDropOverlay({
  sig,
  dispW,
  dispH,
  scale,
  color,
  title,
  onPlace,
  placeK = signPlaceK,
}: {
  sig: SignatureData
  /** Displayed page size at scale=1 (view coords) */
  dispW: number
  dispH: number
  scale: number
  color: [number, number, number]
  title: string
  onPlace: (vx: number, vy: number) => void
  /** Landing-size rule; defaults to signature sizing (image insert passes its own) */
  placeK?: (sig: SignatureData, dispW: number, dispH: number) => number
}): ReactElement {
  const [pt, setPt] = useState<[number, number] | null>(null)
  const k = placeK(sig, dispW, dispH) * scale
  const w = sig.width * k
  const h = sig.height * k
  return (
    <div
      className="pdf-sign-drop"
      title={title}
      onPointerMove={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        setPt([e.clientX - box.left, e.clientY - box.top])
      }}
      onPointerLeave={() => setPt(null)}
      onClick={(e) => {
        const box = e.currentTarget.getBoundingClientRect()
        onPlace((e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
      }}
    >
      {pt && (
        <div
          className="pdf-sign-ghost"
          style={{
            left: Math.min(Math.max(pt[0] - w / 2, 0), Math.max(dispW * scale - w, 0)),
            top: Math.min(Math.max(pt[1] - h / 2, 0), Math.max(dispH * scale - h, 0)),
            width: w,
            height: h,
          }}
        >
          {sig.kind === 'image' ? (
            <img src={`data:image/png;base64,${sig.image}`} alt="" draggable={false} />
          ) : (
            <svg viewBox={`0 0 ${sig.width} ${sig.height}`} preserveAspectRatio="none">
              {sig.paths.map((p, i) => {
                const pts: string[] = []
                for (let j = 0; j < p.length; j += 2) pts.push(`${p[j]},${p[j + 1]}`)
                return (
                  <polyline
                    key={i}
                    points={pts.join(' ')}
                    fill="none"
                    stroke={cssRgb(color)}
                    strokeWidth={2.4}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                )
              })}
            </svg>
          )}
        </div>
      )}
    </div>
  )
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
  const [sidebarW, setSidebarW] = useState(loadSidebarW)
  /** raster width for thumbnails — only updated when a drag ends (re-rastering every frame would jank) */
  const [thumbRasterW, setThumbRasterW] = useState(() => loadSidebarW() - SIDEBAR_CHROME)
  // Re-clamp when the window shrinks (max is 40% of the window), same as slides
  useEffect(() => {
    const onResize = () => setSidebarW((w) => clampSidebarW(w))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])
  /** Drag to resize: width follows the pointer (rAF-throttled); persisted on release */
  const startSidebarResize = (e: React.PointerEvent<HTMLDivElement>) => {
    e.preventDefault()
    const startX = e.clientX
    const startW = sidebarW
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    let w = startW
    let raf = 0
    const onMove = (ev: PointerEvent) => {
      w = clampSidebarW(startW + ev.clientX - startX)
      if (!raf)
        raf = requestAnimationFrame(() => {
          raf = 0
          setSidebarW(w)
        })
    }
    const onUp = () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
      if (raf) cancelAnimationFrame(raf)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setSidebarW(w)
      setThumbRasterW(w - SIDEBAR_CHROME)
      localStorage.setItem(SIDEBAR_W_KEY, String(Math.round(w)))
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }
  const [aiCollapsed, setAiCollapsed] = useState(false)
  /** One-shot prompt pushed by the ribbon AI buttons; the panel auto-runs it (docs preset pattern) */
  const [aiPreset, setAiPreset] = useState<{ text: string; nonce: number } | null>(null)
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home')
  const [spread, setSpread] = useState<1 | 2>(1)
  const [nightMode, setNightMode] = useState(false)
  const [outline, setOutline] = useState<OutlineNode[] | null>(null)
  const [markups, setMarkups] = useState<LocalMarkup[]>([])
  const [drawings, setDrawings] = useState<LocalDrawing[]>([])
  const [drawTool, setDrawTool] = useState<DrawTool | null>(null)
  const [textEdits, setTextEdits] = useState<LocalTextEdit[]>([])
  const [editTextMode, setEditTextMode] = useState(false)
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null)
  /** Hover affordance in edit-text mode: one box over the whole merged line */
  interface LineHover {
    origIdx: number
    box: { left: number; top: number; width: number; height: number }
  }
  const [lineHover, setLineHover] = useState<LineHover | null>(null)
  const lineHoverAnchor = useRef<HTMLElement | null>(null)
  /** Mirror of lineHover for the mousemove handler: state commits lag continuous
      pointer events, so containment must read the just-set box synchronously */
  const lineHoverRef = useRef<LineHover | null>(null)
  const clearLineHover = () => {
    lineHoverAnchor.current = null
    lineHoverRef.current = null
    setLineHover(null)
  }
  /** WPS-style paragraph boxes shown while edit-text mode is on, clustered lazily
      per visible page from the search index (PDF space; cleared on doc reload) */
  const [pageBlocks, setPageBlocks] = useState<Map<number, TextBlock[]>>(new Map())
  const [blockHover, setBlockHover] = useState<{ origIdx: number; idx: number } | null>(null)
  const blockHoverRef = useRef<{ origIdx: number; idx: number } | null>(null)
  const clearBlockHover = () => {
    if (blockHoverRef.current) {
      blockHoverRef.current = null
      setBlockHover(null)
    }
  }
  /** Track the paragraph under the pointer. Runs on every page mousemove (the boxes
      are pointer-events: none so clicks fall through to the text layer), commits
      state only when the hovered block changes */
  const updateBlockHover = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>) => {
    const cur = blockHoverRef.current
    const blocks = pageBlocks.get(origIdx)
    let next: { origIdx: number; idx: number } | null = null
    if (blocks && blocks.length > 0) {
      const pageBox = e.currentTarget.getBoundingClientRect()
      const [px, py] = viewToPdf(
        pageGeom(origIdx),
        (e.clientX - pageBox.left) / scale,
        (e.clientY - pageBox.top) / scale,
      )
      for (let i = 0; i < blocks.length; i++) {
        const r = blocks[i]!.rect
        if (px >= r[0] && px <= r[2] && py >= r[1] && py <= r[3]) {
          next = { origIdx, idx: i }
          break
        }
      }
    }
    if (cur?.origIdx === next?.origIdx && cur?.idx === next?.idx) return
    blockHoverRef.current = next
    setBlockHover(next)
  }
  const updateLineHover = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>) => {
    updateBlockHover(origIdx, e)
    const span = (e.target as HTMLElement).closest('.textLayer span')
    if (!(span instanceof HTMLElement) || !(span.textContent ?? '').trim()) {
      // Within-line gaps hit the textLayer background; keep the affordance while the
      // pointer is still inside the merged box so it doesn't flicker across the line
      const cur = lineHoverRef.current
      if (cur?.origIdx === origIdx) {
        const pageBox = e.currentTarget.getBoundingClientRect()
        const x = e.clientX - pageBox.left
        const y = e.clientY - pageBox.top
        const b = cur.box
        if (x >= b.left && x <= b.left + b.width && y >= b.top && y <= b.top + b.height) return
      }
      clearLineHover()
      return
    }
    if (lineHoverAnchor.current === span) return
    lineHoverAnchor.current = span
    const pageBox = e.currentTarget.getBoundingClientRect()
    const r = groupLineSpans(span).rect
    const next: LineHover = {
      origIdx,
      box: {
        left: r.left - pageBox.left,
        top: r.top - pageBox.top,
        width: r.right - r.left,
        height: r.bottom - r.top,
      },
    }
    lineHoverRef.current = next
    setLineHover(next)
  }
  const [imageEdits, setImageEdits] = useState<LocalImageEdit[]>([])
  /** Latest imageEdits for async callbacks (same rationale as pushUndoRef) */
  const imageEditsRef = useRef(imageEdits)
  imageEditsRef.current = imageEdits
  const [editImageMode, setEditImageMode] = useState(false)
  /** Existing content-stream images per page, listed while edit-image mode is on */
  const [pageImages, setPageImages] = useState<PageImageRef[]>([])
  /** Picked image awaiting click-to-place (same overlay flow as signatures) */
  const [imagePick, setImagePick] = useState<Extract<SignatureData, { kind: 'image' }> | null>(null)
  const imageFileRef = useRef<HTMLInputElement>(null)
  /** Edit-font ids available on this machine (loaded once; empty until then) */
  const [editFonts, setEditFonts] = useState<string[]>([])
  useEffect(() => {
    window.pdfApi
      .listEditFonts()
      .then(setEditFonts)
      .catch(() => {
        /* dropdown simply stays at "original font" */
      })
  }, [])
  /** Select-all runs once per opened draft; refocusing after a style-bar click must keep the caret */
  const draftSelectedRef = useRef(false)
  /** The open draft's textarea: style-bar color clicks read its selection (kept across blur) */
  const draftTaRef = useRef<HTMLTextAreaElement | null>(null)
  /** Colored mirror behind the transparent-text textarea; scroll-synced to it */
  const draftGhostRef = useRef<HTMLDivElement | null>(null)
  const [drawColor, setDrawColor] = useState<[number, number, number]>(DRAW_COLORS[0]!.rgb)
  const [colorOpen, setColorOpen] = useState(false)
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
  const [pendingSign, setPendingSign] = useState<SignatureData | null>(null)
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
  /** Transient message toast (save failures, skipped/rejected text edits) */
  const [notice, setNotice] = useState<string | null>(null)
  const noticeTimerRef = useRef<number | null>(null)
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

  const rows = useMemo(() => spreadRows(visList, spread), [visList, spread])

  /** Visible position → row index */
  const rowOfVis = useCallback((visIdx: number) => rowOfVisIdx(visIdx, spread), [spread])
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

  const loadDoc = useCallback(
    async (path: string, previous: PDFDocumentProxy | null, saved?: SavedSnapshot) => {
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
      if (!saved) {
        setMarkups([])
        setDrawings([])
        setTextEdits([])
        setImageEdits([])
        setTextDraft(null)
        setStampCfg(null)
        setFormEdits(new Map())
        setRotations(new Map())
        setDeleted(new Set())
        setOrder(null)
        setMetadata(null)
      } else {
        // Post-save reload: subtract exactly what the save wrote. Edits made while the
        // write was in flight stay pending, with page indices remapped through the
        // saved deletions/reorder (a page missing from pageMap is gone from the file).
        const remap = saved.pageMap
        setMarkups((prev) =>
          prev.flatMap((mk) => {
            if (saved.markupIds.has(mk.id)) return []
            const ni = remap.get(mk.pageIndex)
            return ni === undefined ? [] : [ni === mk.pageIndex ? mk : { ...mk, pageIndex: ni }]
          }),
        )
        setDrawings((prev) =>
          prev.flatMap((dr) => {
            if (saved.drawingIds.has(dr.id)) return []
            const ni = remap.get(dr.input.pageIndex)
            if (ni === undefined) return []
            return [
              ni === dr.input.pageIndex ? dr : { ...dr, input: { ...dr.input, pageIndex: ni } },
            ]
          }),
        )
        setTextEdits((prev) =>
          prev.flatMap((te) => {
            if (saved.textEditIds.has(te.id)) return []
            const ni = remap.get(te.input.pageIndex)
            if (ni === undefined) return []
            return [
              ni === te.input.pageIndex ? te : { ...te, input: { ...te.input, pageIndex: ni } },
            ]
          }),
        )
        setImageEdits((prev) =>
          prev.flatMap((ie) => {
            if (saved.imageEditIds.has(ie.id)) return []
            const ni = remap.get(ie.input.pageIndex)
            if (ni === undefined) return []
            return [
              ni === ie.input.pageIndex ? ie : { ...ie, input: { ...ie.input, pageIndex: ni } },
            ]
          }),
        )
        setTextDraft((prev) => {
          if (!prev) return null
          const ni = remap.get(prev.origIdx)
          if (ni === undefined) return null
          // A draft re-opened on an edit that just got saved becomes a fresh edit
          const editId =
            prev.editId !== undefined && saved.textEditIds.has(prev.editId)
              ? undefined
              : prev.editId
          return ni === prev.origIdx && editId === prev.editId
            ? prev
            : { ...prev, origIdx: ni, editId }
        })
        // Config-style state: identity compare against the snapshot — unchanged means
        // it is in the file now, a new object means the user changed it during the save
        setStampCfg((prev) => (prev === saved.stampCfg ? null : prev))
        setMetadata((prev) => (prev === saved.metadata ? null : prev))
        setFormEdits((prev) => {
          const next = new Map<string, FormValueInput>()
          for (const [k, v] of prev) if (saved.formEdits.get(k) !== v) next.set(k, v)
          return next
        })
        setRotations((prev) => {
          const next = new Map<number, number>()
          for (const [oldIdx, delta] of prev) {
            const residual = (((delta - (saved.rotations.get(oldIdx) ?? 0)) % 360) + 360) % 360
            const ni = remap.get(oldIdx)
            if (residual !== 0 && ni !== undefined) next.set(ni, residual)
          }
          return next
        })
        setDeleted((prev) => {
          const next = new Set<number>()
          // Saved deletions are absent from pageMap; the rest were deleted mid-save
          for (const oldIdx of prev) {
            const ni = remap.get(oldIdx)
            if (ni !== undefined) next.add(ni)
          }
          return next
        })
        setOrder((prev) => {
          if (!prev) return null
          const mapped = prev.flatMap((o) => {
            const ni = remap.get(o)
            return ni === undefined ? [] : [ni]
          })
          return mapped.every((n, i) => n === i) ? null : mapped
        })
      }
      setFileSize(data.byteLength)
      setSelected(null)
      setDeleteToast(false)
      setUndoStack([])
      setRedoStack([])
      void loaded.getOutline().then(
        (o) => setOutline(o && o.length > 0 ? (o as OutlineNode[]) : null),
        () => setOutline(null),
      )
      // pdfjs-dist 6.x removed PDFDocumentProxy.destroy(); go through the loading task
      if (previous) void previous.loadingTask.destroy()
    },
    [],
  )

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
    textEdits.length > 0 ||
    imageEdits.length > 0 ||
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

  // Existing images are listed while edit-image mode is on; `doc` in the deps refreshes
  // the list after a post-save reload (object rects may have changed on disk)
  useEffect(() => {
    if (!editImageMode || !filePath || !doc) {
      setPageImages([])
      return
    }
    let cancelled = false
    window.pdfApi
      .listPageImages(filePath)
      .then((refs) => {
        if (!cancelled) setPageImages(refs)
      })
      .catch(() => {
        /* hit layer simply stays empty */
      })
    return () => {
      cancelled = true
    }
  }, [editImageMode, filePath, doc])

  // ── Undo/redo: push a full snapshot before each change; consecutive input on the same form field coalesces into one step ──

  const snapshot = (): EditSnapshot => ({
    markups,
    drawings,
    textEdits,
    imageEdits,
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

  /** Latest pushUndo for async callbacks: the AI edit path pushes undo after an awaited
      validation, and the closure it started with may snapshot stale state by then */
  const pushUndoRef = useRef(pushUndo)
  pushUndoRef.current = pushUndo

  const applySnapshot = (s: EditSnapshot) => {
    setMarkups(s.markups)
    setDrawings(s.drawings)
    setTextEdits(s.textEdits)
    setImageEdits(s.imageEdits)
    setTextDraft(null)
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

  /** Paragraph boxes are keyed to the loaded doc; drop them on save-reload */
  useEffect(() => {
    setPageBlocks(new Map())
    clearBlockHover()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc])

  /** Cluster paragraph boxes for pages scrolled into view while edit-text mode is on.
      Reruns after its own setPageBlocks commit and finds nothing missing, so it settles */
  useEffect(() => {
    if (!editTextMode || readOnly || !doc) return
    const missing: number[] = []
    for (const r of visibleRows)
      for (const i of rows[r] ?? []) if (!pageBlocks.has(i)) missing.push(i)
    if (missing.length === 0) return
    const index = getSearchIndex()
    if (!index) return
    let stale = false
    void index.then((entries) => {
      if (stale) return
      setPageBlocks((prev) => {
        const next = new Map(prev)
        for (const i of missing) {
          const entry = entries[i]
          if (!next.has(i)) next.set(i, entry ? groupPageBlocks(entry) : [])
        }
        return next
      })
    })
    return () => {
      stale = true
    }
  }, [editTextMode, readOnly, doc, visibleRows, rows, pageBlocks, getSearchIndex])

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

  /** Shift a drawing by a PDF-space delta (drag-to-move on the page) */
  const moveDrawing = (id: string, dx: number, dy: number) => {
    pushUndo()
    setSelected(null)
    setDrawings((prev) =>
      prev.map((d) => {
        if (d.id !== id) return d
        const input = d.input
        switch (input.kind) {
          case 'ink':
            return {
              ...d,
              input: {
                ...input,
                paths: input.paths.map((p) => p.map((v, i) => (i % 2 === 0 ? v + dx : v + dy))),
              },
            }
          case 'rect':
          case 'ellipse':
          case 'image':
            return {
              ...d,
              input: {
                ...input,
                rect: [
                  input.rect[0] + dx,
                  input.rect[1] + dy,
                  input.rect[2] + dx,
                  input.rect[3] + dy,
                ] as [number, number, number, number],
              },
            }
          case 'line':
          case 'arrow':
            return {
              ...d,
              input: {
                ...input,
                from: [input.from[0] + dx, input.from[1] + dy] as [number, number],
                to: [input.to[0] + dx, input.to[1] + dy] as [number, number],
              },
            }
          default:
            return d
        }
      }),
    )
  }

  /** Replace an image drawing's rect (corner-handle resize) */
  const resizeDrawing = (id: string, rect: [number, number, number, number]) => {
    pushUndo()
    setDrawings((prev) =>
      prev.map((d) =>
        d.id === id && d.input.kind === 'image' ? { ...d, input: { ...d.input, rect } } : d,
      ),
    )
  }

  // ── Text editing (content-stream replacement, applied by the main process at save) ──

  /** Open the paragraph-sized editor over a clustered block: the whole block is the
      edit unit and the commit reflows the text within the block width */
  const startBlockEdit = (origIdx: number, block: TextBlock) => {
    const oldText = joinBlockLines(block.lines.map((l) => l.text))
    if (!oldText.trim()) return
    const rect: [number, number, number, number] = [...block.rect]
    setSelected(null)
    draftSelectedRef.current = false
    setTextDraft({
      origIdx,
      rect,
      oldText,
      fontSize: block.fontSize,
      value: oldText,
      block: {
        leftPt: block.rect[0],
        firstBaseline: block.lines[0]!.y,
        widthPt: block.rect[2] - block.rect[0],
        lineHeight: block.lineHeight,
        align: block.align,
      },
    })
    if (filePath) {
      const probe: TextEditInput = {
        pageIndex: origIdx,
        rect,
        oldText,
        newText: oldText,
        fontSize: block.fontSize,
      }
      void window.pdfApi
        .validateTextEdits({ path: filePath, edits: [probe] })
        .then(([v]) => {
          if (!v?.bounds) return
          setTextDraft((d) =>
            d && !d.editId && d.origIdx === origIdx && d.rect === rect
              ? { ...d, cover: v.bounds }
              : d,
          )
        })
        .catch(() => {
          /* cover falls back to the block rect */
        })
    }
  }

  /** Click on a text-layer span in edit mode → open the floating editor over that run */
  const startTextEdit = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>) => {
    // A plain click inside a multi-line clustered block edits the paragraph
    // (WPS-style); Alt+click keeps the line-level editor as the fallback for
    // clustering misfires. Single-line blocks stay on the line path.
    if (!e.altKey && !readOnly) {
      const blocks = pageBlocks.get(origIdx)
      if (blocks && blocks.length > 0) {
        const pageBox = e.currentTarget.getBoundingClientRect()
        const [px, py] = viewToPdf(
          pageGeom(origIdx),
          (e.clientX - pageBox.left) / scale,
          (e.clientY - pageBox.top) / scale,
        )
        const block = blocks.find(
          (b) => px >= b.rect[0] && px <= b.rect[2] && py >= b.rect[1] && py <= b.rect[3],
        )
        if (block && block.lines.length > 1) {
          e.stopPropagation()
          startBlockEdit(origIdx, block)
          return
        }
      }
    }
    const span = (e.target as HTMLElement).closest('.textLayer span')
    if (!(span instanceof HTMLElement)) return
    if (!(span.textContent ?? '').trim()) return
    // Edit the whole visual line, not the clicked pdf.js run (CJK is often one span
    // per glyph); the save-side matcher aggregates the covered text objects anyway
    const lineGroup = groupLineSpans(span)
    const oldText = lineGroup.text
    if (!oldText.trim()) return
    e.stopPropagation()
    const pageBox = e.currentTarget.getBoundingClientRect()
    const sb = lineGroup.rect
    const geom = pageGeom(origIdx)
    const [ax, ay] = viewToPdf(
      geom,
      (sb.left - pageBox.left) / scale,
      (sb.bottom - pageBox.top) / scale,
    )
    const [bx, by] = viewToPdf(
      geom,
      (sb.right - pageBox.left) / scale,
      (sb.top - pageBox.top) / scale,
    )
    const rect: [number, number, number, number] = [
      Math.min(ax, bx),
      Math.min(ay, by),
      Math.max(ax, bx),
      Math.max(ay, by),
    ]
    const unionH = sb.bottom - sb.top
    const fontSize =
      unionH > 0 ? Math.abs(by - ay) * (lineGroup.fontHeight / unionH) : Math.abs(by - ay)
    setSelected(null)
    draftSelectedRef.current = false
    setTextDraft({ origIdx, rect, oldText, fontSize, value: oldText })
    // The span rect is a font-metric layout box; the run's glyph ink can poke out of it.
    // Fetch the engine's real ink bounds so the editor/preview cover hides the old run fully.
    if (filePath) {
      const probe: TextEditInput = {
        pageIndex: origIdx,
        rect,
        oldText,
        newText: oldText,
        fontSize,
      }
      void window.pdfApi
        .validateTextEdits({ path: filePath, edits: [probe] })
        .then(([v]) => {
          if (!v?.bounds) return
          setTextDraft((d) =>
            d && !d.editId && d.origIdx === origIdx && d.rect === rect
              ? { ...d, cover: v.bounds }
              : d,
          )
        })
        .catch(() => {
          /* cover falls back to the span rect */
        })
    }
  }

  /** Style-bar color pick: a partial textarea selection colors just that range (the
      selection survives the button's focus steal); a collapsed caret or a select-all
      keeps the whole-draft color behavior. `refocus` returns the caret to the textarea
      (skipped for the native color input — its picker panel keeps sending changes). */
  const applyDraftColor = (hex: string, refocus: boolean) => {
    const ta = draftTaRef.current
    setTextDraft((d) => {
      if (!d) return d
      const start = ta?.selectionStart ?? 0
      const end = ta?.selectionEnd ?? 0
      const partial = ta !== null && end > start && !(start === 0 && end >= d.value.length)
      if (!partial) return { ...d, color: hex, charColors: undefined }
      const colors = d.charColors ? [...d.charColors] : Array<string>(d.value.length).fill('')
      for (let i = start; i < end; i++) colors[i] = hex
      return { ...d, charColors: colors }
    })
    if (refocus && ta) {
      const { selectionStart, selectionEnd } = ta
      ta.focus()
      ta.setSelectionRange(selectionStart, selectionEnd)
    }
  }

  /** Fold a floating-editor draft into the pending-edit list; null = nothing changed */
  const mergeTextDraft = (edits: LocalTextEdit[], d: TextDraft): LocalTextEdit[] | null => {
    const existing = d.editId ? edits.find((e) => e.id === d.editId) : undefined
    // Block drafts commit their reflowed form: greedy-wrap each paragraph to the
    // block width in the same face/size the editor previews with, and anchor the
    // rebuilt lines at the block corner with the block's original leading
    let newText = d.value
    let origin: [number, number] | undefined
    let lineLeading: number | undefined
    let lineXOffsets: number[] | undefined
    if (d.block && d.value.trim() !== '') {
      const size = d.size ?? d.fontSize
      const css =
        (d.font ? EDIT_FONT_BY_ID.get(d.font)?.css : undefined) ??
        getComputedStyle(document.body).fontFamily
      const cssStyle = `${d.italic ? 'italic ' : ''}${d.bold ? 'bold' : ''}`.trim()
      lineLeading = d.block.lineHeight * (size / d.fontSize)
      const wrapped = d.value
        .split('\n')
        .flatMap((p) => (p.trim() ? wrapText(p, d.block!.widthPt, size, css, cssStyle) : []))
      newText = wrapped.join('\n')
      origin = [d.block.leftPt, d.block.firstBaseline]
      if (d.block.align !== 'left') {
        lineXOffsets = wrapped.map((l) => {
          const slack = d.block!.widthPt - measurePt(l, size, css, cssStyle)
          return Math.max(0, d.block!.align === 'center' ? slack / 2 : slack)
        })
      }
    }
    // Selection-level colors: draft-space per-char colors carried onto the committed
    // newText (wrapping only rearranges whitespace, so non-ws chars align in order)
    const draftColors = d.charColors?.some((c) => c) ? d.charColors : undefined
    const hexRuns = draftColors
      ? colorsToRuns(
          newText === d.value ? draftColors : mapCharColors(d.value, draftColors, newText),
        )
      : []
    const colorRuns = hexRuns.length
      ? hexRuns.map((r) => ({ start: r.start, end: r.end, color: hexTo255(r.color) }))
      : undefined
    const prevValue = existing ? existing.input.newText : d.oldText
    const cmpValue = existing && d.block ? newText : d.value
    const prevSize = existing?.input.newFontSize
    const prevColor = existing?.input.newColor && rgb255ToHex(existing.input.newColor)
    const prevFont = existing?.input.newFont
    const prevBold = existing?.input.newBold ? true : undefined
    const prevItalic = existing?.input.newItalic ? true : undefined
    const prevRuns = (existing?.input.colorRuns ?? []).map((r) => ({
      start: r.start,
      end: r.end,
      color: rgb255ToHex(r.color),
    }))
    if (
      cmpValue === prevValue &&
      d.size === prevSize &&
      d.color === prevColor &&
      d.font === prevFont &&
      d.bold === prevBold &&
      d.italic === prevItalic &&
      colorRunsEqual(hexRuns, prevRuns)
    )
      return null
    if (
      d.value === d.oldText &&
      d.size === undefined &&
      d.color === undefined &&
      d.font === undefined &&
      d.bold === undefined &&
      d.italic === undefined &&
      hexRuns.length === 0
    ) {
      // Reverted back to the original — the pending edit is moot
      return edits.filter((e) => e.id !== d.editId)
    }
    // A blank replacement would erase the run from the page. Edit mode doesn't offer
    // deletion, so a stray Enter in the emptied box means "never mind", not "wipe it"
    if (d.value.trim() === '') {
      return d.editId ? edits.filter((e) => e.id !== d.editId) : null
    }
    const input: TextEditInput = {
      pageIndex: d.origIdx,
      rect: d.rect,
      oldText: d.oldText,
      newText,
      fontSize: d.fontSize,
      newFontSize: d.size,
      newColor: d.color === undefined ? undefined : hexTo255(d.color),
      colorRuns,
      newFont: d.font,
      newBold: d.bold,
      newItalic: d.italic,
      origin,
      lineLeading,
      lineXOffsets,
      align: d.block && d.block.align !== 'left' ? d.block.align : undefined,
      blockSource: d.block ? d.value : undefined,
    }
    return d.editId
      ? edits.map((e) => (e.id === d.editId ? { ...e, input, cover: d.cover ?? e.cover } : e))
      : [...edits, { id: newId(), input, cover: d.cover }]
  }

  /** Current pending text edits for async callbacks (validation results land after renders) */
  const textEditsRef = useRef(textEdits)
  textEditsRef.current = textEdits

  /** Background dry-run of a just-committed edit against the file. A span that doesn't
      line up with the underlying text objects would otherwise surface only at save time;
      dropping it immediately with a notice beats a save that silently skips it later. */
  const validateTextEdit = (edit: LocalTextEdit) => {
    if (!filePath) return
    void window.pdfApi
      .validateTextEdits({ path: filePath, edits: [edit.input] })
      .then(([v]) => {
        // Stale result: the edit may have been saved or deleted while validation ran
        if (!v || !textEditsRef.current.some((e) => e.id === edit.id)) return
        if (v.reason) {
          setTextEdits((prev) => prev.filter((e) => e.id !== edit.id))
          showNotice(t('textEditNoMatch'))
        } else if (v.bounds) {
          const bounds = v.bounds
          setTextEdits((prev) => prev.map((e) => (e.id === edit.id ? { ...e, cover: bounds } : e)))
        }
      })
      .catch(() => {
        /* best-effort: the save path skips-and-reports unmatched edits anyway */
      })
  }

  /** Close the floating editor and commit its content. Returns the effective edit list
      so save paths can include a just-folded draft that React state hasn't flushed yet. */
  const commitTextDraft = (): LocalTextEdit[] => {
    const d = textDraft
    if (!d) return textEdits
    setTextDraft(null)
    const merged = mergeTextDraft(textEdits, d)
    if (!merged) return textEdits
    pushUndo()
    setTextEdits(merged)
    // New edits append; re-opened ones keep their id
    const committed = d.editId ? merged.find((e) => e.id === d.editId) : merged[merged.length - 1]
    if (committed) validateTextEdit(committed)
    return merged
  }

  const deleteSelected = () => {
    const sel = selected
    if (!sel) return
    pushUndo()
    if (sel.kind === 'markup') setMarkups((prev) => prev.filter((m) => m.id !== sel.id))
    else if (sel.kind === 'drawing') setDrawings((prev) => prev.filter((d) => d.id !== sel.id))
    else if (sel.kind === 'textEdit') setTextEdits((prev) => prev.filter((e) => e.id !== sel.id))
    else if (sel.kind === 'imageEdit') setImageEdits((prev) => prev.filter((e) => e.id !== sel.id))
    else if (sel.kind === 'pageImage')
      // Deleting an untouched existing image = a pending delete op
      setImageEdits((prev) => [
        ...prev,
        {
          id: newId(),
          input: { kind: 'deleteImage', pageIndex: sel.ref.pageIndex, oldRect: sel.ref.rect },
        },
      ])
    else setStampCfg(null)
    setSelected(null)
    // Transient "deleted · undo" toast so the removal is visible and reversible in place
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current)
    setDeleteToast(true)
    toastTimerRef.current = window.setTimeout(() => setDeleteToast(false), 5000)
  }

  /** Show a transient toast; the save-failure badge alone hides the actual reason */
  const showNotice = (msg: string) => {
    if (noticeTimerRef.current !== null) window.clearTimeout(noticeTimerRef.current)
    setNotice(msg)
    noticeTimerRef.current = window.setTimeout(() => setNotice(null), 8000)
  }

  /** Localize known structured main-process errors; other messages pass through raw */
  const friendlySaveError = (error: string): string => {
    const verify = /save-verify-failed pages=([\d,]+)/.exec(error)
    if (verify) return t('saveVerifyFailed', { pages: verify[1]!.split(',').join(', ') })
    return error
  }

  const opFailed = (error: string) => {
    const friendly = friendlySaveError(error)
    setSaveError(friendly)
    setSaveState('error')
    showNotice(`${t('saveFailed')}: ${friendly}`)
  }

  /** Skipped text edits are dropped from the file and from the pending list — surface
      which pages lost an edit instead of silently succeeding */
  const noticeSkippedEdits = (skipped: TextEditFailure[]) => {
    const pages = [...new Set(skipped.map((s) => s.pageIndex + 1))].sort((a, b) => a - b).join(', ')
    showNotice(t('textEditSkipped', { pages }))
  }

  const noticeSkippedImages = (skipped: ImageEditFailure[]) => {
    const pages = [...new Set(skipped.map((s) => s.pageIndex + 1))].sort((a, b) => a - b).join(', ')
    showNotice(t('imageEditSkipped', { pages }))
  }

  /** Pending edits in SavePdfRequest form; shared by in-place Save and Save As */
  const editsPayload = (edits: LocalTextEdit[] = textEdits) => ({
    markups: markups.map(({ id: _id, ...rest }) => rest),
    drawings: drawings.map((d) => d.input),
    textEdits: edits.map((e) => e.input),
    imageEdits: imageEdits.map((e) => e.input),
    stamps: stampCfg ? renderStamps(stampCfg, visList) : [],
    formValues: [...formEdits.values()],
    rotations: [...rotations].map(([pageIndex, delta]) => ({ pageIndex, delta })),
    deletedPages: [...deleted],
    ...(order ? { pageOrder: visList } : {}),
    ...(metadata ? { metadata } : {}),
  })

  /** Resolved when the running save() lands; queued saves and Save As serialize behind it */
  const saveInFlightRef = useRef<Promise<boolean> | null>(null)
  /** Saves requested while another save was writing, drained by an effect after the
      post-save reload has committed. Running the follow-up straight off the completion
      promise would reuse the pre-reload render's closure — dirty still true, the saved
      edits still listed — and write them onto the file a second time. */
  const queuedSavesRef = useRef<{ autosave: boolean; resolve: (ok: boolean) => void }[]>([])

  const save = (autosave = false): Promise<boolean> => {
    // A save is already writing: queue behind it instead of reporting failure — the
    // close prompt's "Save" and ⌘S regularly collide with the blur-triggered autosave
    // (the prompt itself blurs the window). The ref is set synchronously, so this also
    // covers two triggers landing in the same frame, where the saveState snapshot
    // still reads 'idle' for both.
    if (saveInFlightRef.current !== null) {
      return new Promise<boolean>((resolve) => queuedSavesRef.current.push({ autosave, resolve }))
    }
    // Fold an open floating-editor draft in first: keyboard save and autosave can land
    // mid-typing, and the post-save reload closes the editor — without this the
    // in-progress replacement would be silently dropped
    const edits = commitTextDraft()
    const anythingToSave = dirty || edits !== textEdits
    if (!anythingToSave || !filePath) return Promise.resolve(!anythingToSave)
    // An explicit save opts this file into autosave
    if (!autosave) savedOnceRef.current = true
    // What this save writes — the post-save reload subtracts exactly this, keeping
    // any edits the user makes while the write is in flight
    const snapshot: SavedSnapshot = {
      markupIds: new Set(markups.map((mk) => mk.id)),
      drawingIds: new Set(drawings.map((dr) => dr.id)),
      textEditIds: new Set(edits.map((te) => te.id)),
      imageEditIds: new Set(imageEdits.map((ie) => ie.id)),
      stampCfg,
      formEdits,
      rotations,
      metadata,
      pageMap: new Map(visList.map((origIdx, i) => [origIdx, i])),
    }
    const run = (async (): Promise<boolean> => {
      setSaveState('saving')
      const result = await window.pdfApi.save({ path: filePath, ...editsPayload(edits) })
      if (!result.ok) {
        opFailed(result.error)
        return false
      }
      if (result.skippedTextEdits && result.skippedTextEdits.length > 0) {
        noticeSkippedEdits(result.skippedTextEdits)
      }
      if (result.skippedImageEdits && result.skippedImageEdits.length > 0) {
        noticeSkippedImages(result.skippedImageEdits)
      }
      // Reload: changes are in the file now, canvas renders directly, saved pending ops are cleared
      try {
        const el = scrollRef.current
        const scrollTop = el?.scrollTop ?? 0
        await loadDoc(filePath, doc, snapshot)
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

  // Drain queued saves. This effect runs after every commit, so by the time it fires
  // the in-flight save's reload has rendered and `save` reads post-reload state: the
  // follow-up writes only what is still pending (usually nothing) instead of
  // re-applying the previous payload.
  useEffect(() => {
    if (queuedSavesRef.current.length === 0 || saveInFlightRef.current !== null) return
    const queued = queuedSavesRef.current
    queuedSavesRef.current = []
    // One explicit request makes the whole drained batch explicit (autosave opt-in)
    const autosaveOnly = queued.every((q) => q.autosave)
    void save(autosaveOnly).then((ok) => {
      for (const q of queued) q.resolve(ok)
    })
  })

  /**
   * Save As: apply pending edits onto the source bytes and write only to targetPath.
   * The original file stays untouched on disk and the edits stay pending in this tab.
   */
  const saveAsTo = async (targetPath: string): Promise<boolean> => {
    if (!filePath) return false
    // The copy must contain what the user sees, including an open floating-editor draft
    const draftEdits = commitTextDraft()
    // A save already in flight (autosave that started before the dialog opened) lands
    // first. If it succeeded, every edit that was pending is now part of the source
    // bytes, so the copy applies nothing on top — deriving this from the save result
    // (instead of re-reading state) avoids racing React's render of the cleared edits.
    const inFlight = saveInFlightRef.current
    const flushed = inFlight ? await inFlight.catch(() => false) : false
    const edits = flushed
      ? { markups: [], drawings: [], formValues: [], stamps: [], textEdits: [], imageEdits: [] }
      : editsPayload(draftEdits)
    setSaveState('saving')
    const result = await window.pdfApi.save({ path: filePath, targetPath, ...edits })
    if (!result.ok) {
      opFailed(result.error)
      return false
    }
    if (result.skippedTextEdits && result.skippedTextEdits.length > 0) {
      noticeSkippedEdits(result.skippedTextEdits)
    }
    if (result.skippedImageEdits && result.skippedImageEdits.length > 0) {
      noticeSkippedImages(result.skippedImageEdits)
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
      saveInFlightRef.current === null &&
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
    // Image stamps are always drawn upright (both in the overlay and in the saved
    // appearance), so a 90° page turn swaps their displayed width/height. Swap the
    // user-space rect around its center to keep the bitmap's aspect ratio intact.
    setDrawings((prev) =>
      prev.map((d) => {
        if (d.input.kind !== 'image' || d.input.pageIndex !== origIdx) return d
        const [x1, y1, x2, y2] = d.input.rect
        const cx = (x1 + x2) / 2
        const cy = (y1 + y2) / 2
        const hw = (x2 - x1) / 2
        const hh = (y2 - y1) / 2
        return { ...d, input: { ...d.input, rect: [cx - hh, cy - hw, cx + hh, cy + hw] } }
      }),
    )
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

  /** Place signature centered on the click point (view coords, scale=1); sized via signPlaceK
      to match the ghost preview. Points map view→PDF individually so it stays upright on
      rotated pages. */
  const placeSignature = (origIdx: number, vx: number, vy: number) => {
    const sig = pendingSign
    if (!sig) return
    const geom = pageGeom(origIdx)
    const disp = geomDispSize(geom)
    const k = signPlaceK(sig, disp.width, disp.height)
    const targetW = sig.width * k
    const targetH = sig.height * k
    const left = Math.min(Math.max(vx - targetW / 2, 0), Math.max(disp.width - targetW, 0))
    const top = Math.min(Math.max(vy - targetH / 2, 0), Math.max(disp.height - targetH, 0))
    pushUndo()
    if (sig.kind === 'image') {
      const [ax, ay] = viewToPdf(geom, left, top)
      const [bx, by] = viewToPdf(geom, left + targetW, top + targetH)
      const rect: [number, number, number, number] = [
        Math.min(ax, bx),
        Math.min(ay, by),
        Math.max(ax, bx),
        Math.max(ay, by),
      ]
      setDrawings((prev) => [
        ...prev,
        { id: newId(), input: { kind: 'image', pageIndex: origIdx, image: sig.image, rect } },
      ])
    } else {
      const paths = sig.paths.map((p) => {
        const out: number[] = []
        for (let i = 0; i < p.length; i += 2) {
          out.push(...viewToPdf(geom, left + p[i]! * k, top + p[i + 1]! * k))
        }
        return out
      })
      setDrawings((prev) => [
        ...prev,
        {
          id: newId(),
          input: { kind: 'ink', pageIndex: origIdx, color: drawColor, width: 1.6, paths },
        },
      ])
    }
    setPendingSign(null)
  }

  // ── Image editing (content-stream image ops, applied by the main process at save) ──

  /** Ribbon button → file picker; the picked image then enters click-to-place mode */
  const pickInsertImage = () => {
    setEditTextMode(false)
    setTextDraft(null)
    setDrawTool(null)
    setPendingSign(null)
    replaceTargetRef.current = null
    imageFileRef.current?.click()
  }

  const onImageFilePicked = async (file: File) => {
    const canvas = await fileToCanvas(file, 2400)
    const base64 = canvas?.toDataURL('image/png').split(',')[1]
    if (!canvas || !base64) return
    const target = replaceTargetRef.current
    if (target) {
      replaceTargetRef.current = null
      replaceExisting(target, base64)
      return
    }
    setImagePick({ kind: 'image', image: base64, width: canvas.width, height: canvas.height })
  }

  /** Drop the picked image centered on the click point, into the text-below band by default */
  const placeImage = (origIdx: number, vx: number, vy: number) => {
    const pick = imagePick
    if (!pick) return
    const geom = pageGeom(origIdx)
    const disp = geomDispSize(geom)
    const k = imagePlaceK(pick, disp.width, disp.height)
    const w = pick.width * k
    const h = pick.height * k
    const left = Math.min(Math.max(vx - w / 2, 0), Math.max(disp.width - w, 0))
    const top = Math.min(Math.max(vy - h / 2, 0), Math.max(disp.height - h, 0))
    const [ax, ay] = viewToPdf(geom, left, top)
    const [bx, by] = viewToPdf(geom, left + w, top + h)
    pushUndo()
    setImageEdits((prev) => [
      ...prev,
      {
        id: newId(),
        input: {
          kind: 'insertImage',
          pageIndex: origIdx,
          image: pick.image,
          rect: [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)],
          layer: 'belowText',
          rotate: ((geom.rot % 360) + 360) % 360,
        },
      },
    ])
    setImagePick(null)
  }

  /** Committed move/resize of a pending image op */
  const updateImageEditRect = (id: string, rect: [number, number, number, number]) => {
    pushUndo()
    setSelected(null)
    setImageEdits((prev) =>
      prev.map((e) =>
        e.id === id && e.input.kind !== 'deleteImage' ? { ...e, input: { ...e.input, rect } } : e,
      ),
    )
  }

  /** Prefetched pixels of untouched existing images (keyed pageIndex:rectKey); fetched on
      select/drag-start so the picture can follow the hand before any op exists */
  const [existingPngs, setExistingPngs] = useState<Map<string, string>>(new Map())
  const existingPngFetches = useRef(new Set<string>())
  /** Bumped whenever the cache is dropped; a late IPC response started against a
      previous doc must not repopulate the fresh cache (rect keys can collide) */
  const existingPngEpoch = useRef(0)

  const prefetchExistingPng = (ref: PageImageRef) => {
    const key = `${ref.pageIndex}:${imageRectKey(ref.rect)}`
    if (existingPngFetches.current.has(key)) return
    existingPngFetches.current.add(key)
    const epoch = existingPngEpoch.current
    void window.pdfApi
      .pageImagePng({ path: filePath, pageIndex: ref.pageIndex, rect: ref.rect })
      .then((png) => {
        if (existingPngEpoch.current !== epoch) return
        if (png) setExistingPngs((prev) => new Map(prev).set(key, png))
      })
      .catch(() => {
        if (existingPngEpoch.current === epoch) existingPngFetches.current.delete(key)
      })
  }

  // Image rects change identity when the file is saved/reloaded; drop the cache
  useEffect(() => {
    existingPngEpoch.current++
    setExistingPngs(new Map())
    existingPngFetches.current.clear()
  }, [doc])

  /** First touch of an existing image (drag/resize/layer/rotate) becomes a pending transform
      op; its rendered pixels come from the prefetch cache or are fetched for the ghost preview */
  const transformExisting = (
    ref: PageImageRef,
    rect: [number, number, number, number],
    layer?: ImageLayer,
    quarterTurns?: number,
  ) => {
    pushUndo()
    setSelected(null)
    const id = newId()
    const cached = existingPngs.get(`${ref.pageIndex}:${imageRectKey(ref.rect)}`) ?? null
    setImageEdits((prev) => [
      ...prev,
      {
        id,
        input: {
          kind: 'transformImage',
          pageIndex: ref.pageIndex,
          oldRect: ref.rect,
          rect,
          ...(layer ? { layer } : {}),
          ...(quarterTurns ? { quarterTurns } : {}),
        },
        png: cached,
        origAbove: ref.aboveText,
      },
    ])
    if (cached) return
    void window.pdfApi
      .pageImagePng({ path: filePath, pageIndex: ref.pageIndex, rect: ref.rect })
      .then((png) => {
        if (png) setImageEdits((prev) => prev.map((e) => (e.id === id ? { ...e, png } : e)))
      })
      .catch(() => {
        /* ghost stays a dashed box */
      })
  }

  /** The rect's footprint after an odd quarter turn about its center (w/h swap) */
  const rotatedRect = (r: readonly number[]): [number, number, number, number] => {
    const cx = (r[0]! + r[2]!) / 2
    const cy = (r[1]! + r[3]!) / 2
    const w = r[2]! - r[0]!
    const h = r[3]! - r[1]!
    return [cx - h / 2, cy - w / 2, cx + h / 2, cy + w / 2]
  }

  /** Rotate PNG pixels by a quarter turn (fresh inserts carry rotation baked into the bytes) */
  const rotatePng90 = (b64: string, cw: boolean): Promise<string | null> =>
    new Promise((resolve) => {
      const img = new Image()
      img.onload = () => {
        const c = document.createElement('canvas')
        c.width = img.height
        c.height = img.width
        const ctx = c.getContext('2d')
        if (!ctx) return resolve(null)
        ctx.translate(c.width / 2, c.height / 2)
        ctx.rotate(((cw ? 90 : -90) * Math.PI) / 180)
        ctx.drawImage(img, -img.width / 2, -img.height / 2)
        resolve(c.toDataURL('image/png').split(',')[1] ?? null)
      }
      img.onerror = () => resolve(null)
      img.src = `data:image/png;base64,${b64}`
    })

  /** Rotate the selected image a quarter turn (screen-clockwise when dir = 1) */
  const rotateSelected = (dir: 1 | -1) => {
    const sel = selected
    if (!sel) return
    const turn = dir === 1 ? 1 : 3
    if (sel.kind === 'pageImage') {
      prefetchExistingPng(sel.ref)
      transformExisting(sel.ref, rotatedRect(sel.ref.rect), undefined, turn)
      return
    }
    if (sel.kind !== 'imageEdit') return
    const edit = imageEdits.find((e) => e.id === sel.id)
    if (!edit || edit.input.kind === 'deleteImage') return
    setSelected(null)
    if (edit.input.kind === 'insertImage') {
      const { image } = edit.input
      void rotatePng90(image, dir === 1).then((rotated) => {
        if (!rotated) return
        // The canvas turn is async: snapshot via the ref (the closed-over pushUndo
        // would capture click-time state) and rotate the element's CURRENT rect —
        // a concurrent move/resize must not be overwritten. The bytes guard drops
        // a rotation that lost a race (edit removed, or another turn landed first)
        // BEFORE pushing undo, so ⌘Z never records a no-op step.
        const target = imageEditsRef.current.find((e) => e.id === sel.id)
        if (!target || target.input.kind !== 'insertImage' || target.input.image !== image) return
        pushUndoRef.current()
        setImageEdits((prev) =>
          prev.map((e) =>
            e.id === sel.id && e.input.kind === 'insertImage' && e.input.image === image
              ? { ...e, input: { ...e.input, image: rotated, rect: rotatedRect(e.input.rect) } }
              : e,
          ),
        )
      })
      return
    }
    pushUndo()
    setImageEdits((prev) =>
      prev.map((e) =>
        e.id === sel.id && (e.input.kind === 'transformImage' || e.input.kind === 'replaceImage')
          ? {
              ...e,
              input: {
                ...e.input,
                quarterTurns: ((e.input.quarterTurns ?? 0) + turn) % 4,
                rect: rotatedRect(e.input.rect),
              },
            }
          : e,
      ),
    )
  }

  /** Queue an in-place pixel swap of an existing image (footprint/z-order kept) */
  const replaceExisting = (ref: PageImageRef, png: string) => {
    pushUndo()
    setSelected(null)
    setImageEdits((prev) => [
      ...prev,
      {
        id: newId(),
        input: {
          kind: 'replaceImage',
          pageIndex: ref.pageIndex,
          oldRect: ref.rect,
          rect: ref.rect,
          image: png,
        },
        origAbove: ref.aboveText,
      },
    ])
  }

  /** Replace flow: the bubble button stashes the target, then reuses the insert file input */
  const replaceTargetRef = useRef<PageImageRef | null>(null)
  const startReplaceImage = () => {
    const sel = selected
    if (sel?.kind !== 'pageImage') return
    replaceTargetRef.current = sel.ref
    setSelected(null)
    imageFileRef.current?.click()
  }

  /** Current z-band of the selected image thing (labels the popup toggle) */
  const selectedImageLayer = (): ImageLayer | null => {
    const sel = selected
    if (sel?.kind === 'pageImage') return sel.ref.aboveText ? 'aboveText' : 'belowText'
    if (sel?.kind !== 'imageEdit') return null
    const e = imageEdits.find((x) => x.id === sel.id)
    if (!e || e.input.kind === 'deleteImage') return null
    if (e.input.kind === 'insertImage') return e.input.layer
    return e.input.layer ?? (e.origAbove ? 'aboveText' : 'belowText')
  }

  const toggleImageLayer = () => {
    const sel = selected
    const cur = selectedImageLayer()
    if (!sel || !cur) return
    const next: ImageLayer = cur === 'aboveText' ? 'belowText' : 'aboveText'
    if (sel.kind === 'pageImage') {
      transformExisting(sel.ref, sel.ref.rect, next)
    } else if (sel.kind === 'imageEdit') {
      pushUndo()
      setImageEdits((prev) =>
        prev.map((e) =>
          e.id === sel.id && e.input.kind !== 'deleteImage'
            ? { ...e, input: { ...e.input, layer: next } }
            : e,
        ),
      )
    }
    setSelected(null)
  }

  /** Existing images already claimed by a pending op are hidden from the hit layer */
  const claimedImageKeys = useMemo(
    () =>
      new Set(
        imageEdits.flatMap((e) =>
          e.input.kind === 'insertImage'
            ? []
            : [`${e.input.pageIndex}:${imageRectKey(e.input.oldRect)}`],
        ),
      ),
    [imageEdits],
  )

  // ── Live page preview: pdfium re-renders the touched region without the moved/
  // resized/deleted images, so the original vanishes immediately instead of at save ──

  /** Pages with pending transform/delete ops → the original rects to erase */
  const livePreviewRects = useMemo(() => {
    const map = new Map<number, [number, number, number, number][]>()
    for (const e of imageEdits) {
      if (e.input.kind !== 'insertImage') {
        map.set(e.input.pageIndex, [...(map.get(e.input.pageIndex) ?? []), e.input.oldRect])
      }
    }
    return map
  }, [imageEdits])

  const [livePreview, setLivePreview] = useState<
    Map<number, { png: string; clip: { x: number; y: number; width: number; height: number } }>
  >(new Map())
  /** Last requested render key per page; skips redundant IPC round-trips */
  const livePreviewKeys = useRef(new Map<number, string>())

  useEffect(() => {
    // Drop previews for pages whose ops are gone (undo / save reload)
    for (const p of [...livePreviewKeys.current.keys()]) {
      if (!livePreviewRects.has(p)) livePreviewKeys.current.delete(p)
    }
    setLivePreview((prev) => {
      if (![...prev.keys()].some((p) => !livePreviewRects.has(p))) return prev
      const next = new Map(prev)
      for (const p of [...next.keys()]) if (!livePreviewRects.has(p)) next.delete(p)
      return next
    })
    for (const [pageIndex, rects] of livePreviewRects) {
      const geom = pageGeom(pageIndex)
      const disp = geomDispSize(geom)
      // Union of the erased rects in display coords, padded and clamped to the page
      let x1 = Infinity
      let y1 = Infinity
      let x2 = -Infinity
      let y2 = -Infinity
      for (const r of rects) {
        const b = pdfRectToCss(geom, r, 1)
        x1 = Math.min(x1, b.left)
        y1 = Math.min(y1, b.top)
        x2 = Math.max(x2, b.left + b.width)
        y2 = Math.max(y2, b.top + b.height)
      }
      const pad = 2
      x1 = Math.max(0, x1 - pad)
      y1 = Math.max(0, y1 - pad)
      x2 = Math.min(disp.width, x2 + pad)
      y2 = Math.min(disp.height, y2 + pad)
      if (x2 - x1 <= 0 || y2 - y1 <= 0) continue
      const clip = { x: x1, y: y1, width: x2 - x1, height: y2 - y1 }
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      const pxWidth = Math.min(Math.ceil(clip.width * scale * dpr), 2800)
      const rotIdx = ((Math.round(rotDelta(pageIndex) / 90) % 4) + 4) % 4
      const key = `${rects.map(imageRectKey).join(';')}|${pxWidth}|${rotIdx}`
      if (livePreviewKeys.current.get(pageIndex) === key) continue
      livePreviewKeys.current.set(pageIndex, key)
      void window.pdfApi
        .pagePreviewPng({
          path: filePath,
          pageIndex,
          excludeRects: rects,
          clip,
          pxWidth,
          rotate: rotIdx,
        })
        .then((png) => {
          // Stale guard: a newer request for this page may have superseded this one
          // while the render ran — its result must not be overwritten by ours
          if (livePreviewKeys.current.get(pageIndex) !== key) return
          if (png) setLivePreview((prev) => new Map(prev).set(pageIndex, { png, clip }))
        })
        .catch(() => {
          // Only clear the key if it is still ours; deleting a newer in-flight key
          // would let its (valid) result be treated as stale and the page re-request
          if (livePreviewKeys.current.get(pageIndex) === key) {
            livePreviewKeys.current.delete(pageIndex)
          }
        })
    }
  }, [livePreviewRects, scale, pageGeom, rotDelta, filePath])

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
          void pdoc.loadingTask.destroy()
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
    editText: async (input) => {
      let cover: [number, number, number, number] | undefined
      if (filePath) {
        try {
          const [v] = await window.pdfApi.validateTextEdits({ path: filePath, edits: [input] })
          if (v?.reason) return v.reason
          cover = v?.bounds
        } catch {
          /* best-effort: the save path skips-and-reports unmatched edits anyway */
        }
      }
      pushUndoRef.current()
      setTextEdits((prev) => [...prev, { id: newId(), input, cover }])
      return null
    },
    editFonts: () => editFonts,
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
    pageGeom: (origIdx) => (sizes[origIdx] ? pageGeom(origIdx) : null),
    listImages: () => (filePath ? window.pdfApi.listPageImages(filePath) : Promise.resolve([])),
    isImageClaimed: (ref) => claimedImageKeys.has(`${ref.pageIndex}:${imageRectKey(ref.rect)}`),
    insertImage: (origIdx, png, rect, layer) => {
      pushUndoRef.current()
      setImageEdits((prev) => [
        ...prev,
        {
          id: newId(),
          input: {
            kind: 'insertImage',
            pageIndex: origIdx,
            image: png,
            rect,
            layer,
            rotate: ((pageGeom(origIdx).rot % 360) + 360) % 360,
          },
        },
      ])
    },
    transformImage: (ref, rect, layer, quarterTurns) =>
      transformExisting(ref, rect, layer, quarterTurns),
    replaceImage: (ref, png) => replaceExisting(ref, png),
    deleteImage: (ref) => {
      pushUndoRef.current()
      setImageEdits((prev) => [
        ...prev,
        {
          id: newId(),
          input: { kind: 'deleteImage', pageIndex: ref.pageIndex, oldRect: ref.rect },
        },
      ])
    },
    searchImages: (query, maxResults) => window.pdfApi.imageSearch(query, maxResults),
    generateImage: (op) => window.pdfApi.generateImage(op),
    fetchImage: async (url) => {
      const fetched = await window.pdfApi.fetchImage(url)
      if (!fetched) return null
      try {
        const bytes = Uint8Array.from(atob(fetched.base64), (c) => c.charCodeAt(0))
        const canvas = await fileToCanvas(
          new File([bytes], 'ai-image', { type: fetched.mime }),
          2400,
        )
        const png = canvas?.toDataURL('image/png').split(',')[1]
        return canvas && png ? { png, width: canvas.width, height: canvas.height } : null
      } catch {
        return null
      }
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

  // Clicking elsewhere closes the draw-color palette
  useEffect(() => {
    if (!colorOpen) return
    const close = (e: PointerEvent) => {
      if (!(e.target as Element | null)?.closest?.('.rb-drop-wrap')) setColorOpen(false)
    }
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [colorOpen])

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
        if (textDraft) setTextDraft(null)
        else if (imagePick) setImagePick(null)
        else if (editTextMode) setEditTextMode(false)
        else if (editImageMode) setEditImageMode(false)
        else if (pendingSign) setPendingSign(null)
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
      const inThumbs = !!thumbsRef.current?.contains(document.activeElement)
      const action = navAction(e.key, inThumbs)
      if (!action) return
      e.preventDefault()
      switch (action.type) {
        case 'scrollViewport':
          el.scrollTop += action.dir * (el.clientHeight - 40)
          break
        case 'scrollEdge':
          el.scrollTop = action.edge === 'top' ? 0 : el.scrollHeight
          break
        case 'scrollBy':
          el.scrollTop += action.delta
          break
        case 'stepPage': {
          const target = stepPage(visList, spread, currentPage, action.dir)
          scrollToPage(target)
          if (inThumbs) {
            const thumbEl = thumbsRef.current?.querySelector<HTMLElement>(
              `[data-idx="${target - 1}"]`,
            )
            thumbEl?.focus({ preventScroll: true })
            thumbEl?.scrollIntoView({ block: 'nearest' })
          }
          break
        }
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

  /** Ribbon AI buttons: expand the dock and auto-run the prompt in the assistant */
  const runAiPreset = (text: string): void => {
    setAiCollapsed(false)
    setAiPreset({ text, nonce: Date.now() })
  }

  // ── shared ribbon groups (rendered on more than one tab) ──
  // mousedown preventDefault: the browser clears the text selection the instant the button is pressed, so applyMarkup would lose it
  const markupGroup = (
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
  )

  const pageZoomGroup = (
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
  )

  const searchBtn = (
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
  )

  const editTextBtn = (
    <button
      className={`rb-big${editTextMode ? ' active' : ''}`}
      disabled={readOnly}
      title={t('editTextHint')}
      onClick={() => {
        setTextDraft(null)
        setDrawTool(null)
        setPendingSign(null)
        setImagePick(null)
        setEditImageMode(false)
        setEditTextMode((v) => !v)
      }}
    >
      <span className="rb-big-icon">
        <IconEditText />
      </span>
      {t('editText')}
    </button>
  )

  const viewNavGroup = (
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
        {searchBtn}
        <button
          className={`rb-big${spread === 2 ? ' active' : ''}`}
          title={spread === 2 ? t('singlePage') : t('twoPage')}
          onClick={() => setSpread((v) => (v === 1 ? 2 : 1))}
        >
          <span className="rb-big-icon">{spread === 2 ? <IconSinglePage /> : <IconSpread />}</span>
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
  )

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
          <span className="qa-sep" />
          {RIBBON_TABS.map(({ id, labelKey }) => (
            <button
              key={id}
              className={`ribbon-tab${ribbonTab === id ? ' active' : ''}`}
              onClick={() => setRibbonTab(id)}
            >
              {t(labelKey)}
            </button>
          ))}
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
          {ribbonTab === 'home' && (
            <>
              {/* ---- Genspark AI (first slot: entry + one-click AI actions, docs parity) ---- */}
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big ai-entry${aiCollapsed ? '' : ' active'}`}
                    title={t('aiOpenAssistant')}
                    onClick={() => setAiCollapsed((v) => !v)}
                  >
                    <span className="rb-big-icon">
                      <GensparkMark size={26} />
                    </span>
                    <span>Genspark AI</span>
                  </button>
                  <button
                    className="rb-big ai-entry"
                    title={t('aiQuickSummaryPrompt')}
                    onClick={() => runAiPreset(t('aiQuickSummaryPrompt'))}
                  >
                    <span className="rb-big-icon">
                      <span className="ai-feature-icon" aria-hidden="true">
                        <IconAiSummarize />
                      </span>
                    </span>
                    <span>{t('aiSummarizeBtn')}</span>
                  </button>
                  <button
                    className="rb-big ai-entry"
                    title={t('aiQuickKeyPointsPrompt')}
                    onClick={() => runAiPreset(t('aiQuickKeyPointsPrompt'))}
                  >
                    <span className="rb-big-icon">
                      <span className="ai-feature-icon" aria-hidden="true">
                        <IconAiKeyPoints />
                      </span>
                    </span>
                    <span>{t('aiKeyPointsBtn')}</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              {markupGroup}
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  {searchBtn}
                  {editTextBtn}
                </div>
              </div>
              <div className="ribbon-sep" />
              {pageZoomGroup}
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
                    title={t('propsTitle')}
                    onClick={() => setPropsDlg(true)}
                  >
                    <span className="rb-big-icon">
                      <IconProps />
                    </span>
                    {t('props')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'annotate' && (
            <>
              {markupGroup}
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  {DRAW_TOOLS.map(({ tool, icon: DrawIcon, key }) => (
                    <button
                      key={tool}
                      className={`rb-big${drawTool === tool ? ' active' : ''}`}
                      disabled={readOnly}
                      title={t(key)}
                      onClick={() => {
                        setEditTextMode(false)
                        setTextDraft(null)
                        setImagePick(null)
                        setEditImageMode(false)
                        setDrawTool((v) => (v === tool ? null : tool))
                      }}
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
                    onClick={() => {
                      setImagePick(null)
                      setEditImageMode(false)
                      if (pendingSign) setPendingSign(null)
                      else setSignDlg(true)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconSign />
                    </span>
                    {t('sign')}
                  </button>
                  <div className="rb-drop-wrap">
                    <button
                      className={`rb-big${colorOpen ? ' active' : ''}`}
                      disabled={readOnly}
                      title={t('drawColor')}
                      onClick={() => setColorOpen((v) => !v)}
                    >
                      <span className="rb-big-icon">
                        <span className="rb-big-icon-colored">
                          <IconDrawColor />
                          <span
                            className="rb-color-bar"
                            style={{ background: cssRgb(drawColor) }}
                          />
                        </span>
                        <RbCaret />
                      </span>
                      {t('drawColor')}
                    </button>
                    {colorOpen && (
                      <div className="rb-drop rb-color-grid">
                        {DRAW_COLORS.map((c) => (
                          <button
                            key={c.name}
                            className={`rb-swatch${cssRgb(drawColor) === cssRgb(c.rgb) ? ' active' : ''}`}
                            style={{ background: cssRgb(c.rgb) }}
                            title={c.name}
                            onClick={() => {
                              setDrawColor(c.rgb)
                              setColorOpen(false)
                            }}
                          />
                        ))}
                        <label className="rb-color-more" title={t('drawColor')}>
                          <input
                            type="color"
                            value={rgbToHex(drawColor)}
                            onChange={(e) => setDrawColor(hexToRgb(e.target.value))}
                          />
                          {t('moreColors')}
                        </label>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'edit' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  {editTextBtn}
                  <button
                    className={`rb-big${imagePick ? ' active' : ''}`}
                    disabled={readOnly}
                    title={t('insertImageHint')}
                    onClick={() => (imagePick ? setImagePick(null) : pickInsertImage())}
                  >
                    <span className="rb-big-icon">
                      <IconInsertImage />
                    </span>
                    {t('insertImage')}
                  </button>
                  <button
                    className={`rb-big${editImageMode ? ' active' : ''}`}
                    disabled={readOnly}
                    title={t('editImageHint')}
                    onClick={() => {
                      setEditTextMode(false)
                      setTextDraft(null)
                      setDrawTool(null)
                      setPendingSign(null)
                      setImagePick(null)
                      setEditImageMode((v) => !v)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconEditImage />
                    </span>
                    {t('editImage')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
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
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'page' && (
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
          )}
          {ribbonTab === 'view' && (
            <>
              {viewNavGroup}
              <div className="ribbon-sep" />
              {pageZoomGroup}
            </>
          )}
        </div>
      </div>
      <input
        ref={imageFileRef}
        type="file"
        accept="image/png,image/jpeg,image/webp,image/gif,image/bmp"
        style={{ display: 'none' }}
        onChange={(e) => {
          const f = e.target.files?.[0]
          e.target.value = ''
          if (f) void onImageFilePicked(f)
        }}
      />
      <div className="app-main">
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
          <AiPanel api={aiApi} preset={aiPreset} onCollapse={() => setAiCollapsed(true)} />
        </div>
        <div className="app-content">
          <div className="pdf-body">
            {sidebar === 'outline' && outline && (
              <div className="pdf-thumbs pdf-outline-pane" style={{ width: sidebarW }}>
                <OutlinePanel outline={outline} onGoToDest={(dest) => void goToDest(dest)} />
              </div>
            )}
            {sidebar === 'thumbs' && (
              <div ref={thumbsRef} className="pdf-thumbs" style={{ width: sidebarW }}>
                {visList.map((origIdx, v) => {
                  const size = dispSize(origIdx)
                  return (
                    <div
                      key={origIdx}
                      ref={setThumbRef(v)}
                      data-idx={v}
                      tabIndex={-1}
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
                        style={{ aspectRatio: `${size.width} / ${size.height}` }}
                      >
                        <PdfThumb
                          doc={doc}
                          pageNo={origIdx + 1}
                          rotationDelta={rotDelta(origIdx)}
                          visible={visibleThumbs.has(v)}
                          rasterW={thumbRasterW}
                        />
                      </div>
                      <span className="pdf-thumb-no">{v + 1}</span>
                    </div>
                  )
                })}
              </div>
            )}
            {(sidebar === 'thumbs' || (sidebar === 'outline' && !!outline)) && (
              <div className="pdf-side-resizer" onPointerDown={startSidebarResize} />
            )}
            <div
              ref={scrollRef}
              className={`pdf-scroll${drawTool ? ' pdf-drawing' : ''}${nightMode ? ' pdf-night' : ''}`}
              onScroll={() => {
                handleScroll()
                setSelPopup(null)
                setSelected(null)
                clearLineHover()
                clearBlockHover()
              }}
              onMouseUp={drawTool ? undefined : handleMouseUp}
              onClick={(e) => {
                // Clicking anywhere that isn't an annotation clears the selection
                if (
                  !(e.target as Element).closest?.(
                    '.pdf-markup, .pdf-draw-shape, .pdf-note-pin, .pdf-stamp-preview, .pdf-textedit-preview, .pdf-textedit-input, .pdf-imgedit-layer, .pdf-imgedit-under',
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
                        className={`pdf-page${editTextMode && !readOnly ? ' pdf-editing-text' : ''}`}
                        style={
                          {
                            width: Math.floor(size.width * scale),
                            height: Math.floor(size.height * scale),
                            '--scale-factor': scale,
                          } as CSSProperties
                        }
                        onClick={
                          editTextMode && !readOnly ? (e) => startTextEdit(origIdx, e) : undefined
                        }
                        onMouseMove={
                          // move, not over: leaving the hover box across the static textLayer
                          // background fires no over events; the handler early-returns while
                          // the anchor span is unchanged so per-move cost is one closest()
                          editTextMode && !readOnly ? (e) => updateLineHover(origIdx, e) : undefined
                        }
                        onMouseLeave={
                          editTextMode && !readOnly
                            ? () => {
                                clearLineHover()
                                clearBlockHover()
                              }
                            : undefined
                        }
                      >
                        <PdfPage
                          doc={doc}
                          pageNo={origIdx + 1}
                          scale={scale}
                          rotationDelta={rotDelta(origIdx)}
                          visible={rowVisible}
                        />
                        {livePreview.has(origIdx) &&
                          (() => {
                            const lp = livePreview.get(origIdx)!
                            return (
                              <img
                                className="pdf-page-livepreview"
                                src={`data:image/png;base64,${lp.png}`}
                                alt=""
                                style={{
                                  left: lp.clip.x * scale,
                                  top: lp.clip.y * scale,
                                  width: lp.clip.width * scale,
                                  height: lp.clip.height * scale,
                                }}
                              />
                            )
                          })()}
                        {pendingSign && (
                          <SignDropOverlay
                            sig={pendingSign}
                            dispW={geomDispSize(geom).width}
                            dispH={geomDispSize(geom).height}
                            scale={scale}
                            color={drawColor}
                            title={t('signHint')}
                            onPlace={(vx, vy) => placeSignature(origIdx, vx, vy)}
                          />
                        )}
                        {imagePick && (
                          <SignDropOverlay
                            sig={imagePick}
                            dispW={geomDispSize(geom).width}
                            dispH={geomDispSize(geom).height}
                            scale={scale}
                            color={drawColor}
                            title={t('imagePlaceHint')}
                            onPlace={(vx, vy) => placeImage(origIdx, vx, vy)}
                            placeK={imagePlaceK}
                          />
                        )}
                        {/* Paragraph boxes (WPS-style): every text block outlined while
                            edit-text mode is on; hovered one highlighted, all dimmed
                            while the floating editor is open */}
                        {editTextMode &&
                          !readOnly &&
                          rowVisible &&
                          (pageBlocks.get(origIdx) ?? []).map((b, i) => (
                            <div
                              key={i}
                              className={`pdf-textblock-box${
                                blockHover?.origIdx === origIdx && blockHover.idx === i
                                  ? ' is-hover'
                                  : ''
                              }${textDraft ? ' is-faded' : ''}`}
                              style={pdfRectToCss(geom, b.rect, scale)}
                            />
                          ))}
                        {editTextMode && !readOnly && lineHover?.origIdx === origIdx && (
                          <div className="pdf-textline-hover" style={lineHover.box} />
                        )}
                        {rowVisible && (
                          <>
                            {/* Pending text edits: cover the original run and preview the replacement */}
                            {textEdits
                              .filter((te) => te.input.pageIndex === origIdx)
                              .map((te) => {
                                const fs =
                                  (te.input.newFontSize ?? te.input.fontSize) * scale * 0.92
                                const lineCount = te.input.newText.split('\n').length
                                const leadPx = te.input.lineLeading
                                  ? te.input.lineLeading * scale
                                  : fs * 1.2
                                const style: CSSProperties = {
                                  ...pdfRectToCss(geom, te.input.rect, scale),
                                  fontSize: fs,
                                  ...(te.input.lineLeading ? { lineHeight: `${leadPx}px` } : {}),
                                }
                                if (te.input.newColor) {
                                  style.color = `rgb(${te.input.newColor.join(', ')})`
                                }
                                if (te.input.newFont) {
                                  style.fontFamily = EDIT_FONT_BY_ID.get(te.input.newFont)?.css
                                }
                                if (te.input.newBold) style.fontWeight = 700
                                if (te.input.newItalic) style.fontStyle = 'italic'
                                if (lineCount > 1) {
                                  // Grow below the original rect, same leading the engine writes
                                  // (block edits carry the paragraph's own leading)
                                  style.height = lineCount * leadPx
                                  style.lineHeight = te.input.lineLeading ? `${leadPx}px` : 1.2
                                  style.alignItems = 'flex-start'
                                }
                                // The rebuilt run grows right past the original rect when the
                                // replacement is longer; the preview must too, or the extra
                                // characters look cut off until the save (overflow: hidden)
                                const previewFont = `${te.input.newItalic ? 'italic ' : ''}${
                                  te.input.newBold ? 'bold ' : ''
                                }${fs}px ${
                                  (te.input.newFont &&
                                    EDIT_FONT_BY_ID.get(te.input.newFont)?.css) ||
                                  getComputedStyle(document.body).fontFamily
                                }`
                                const widest = Math.max(
                                  ...te.input.newText
                                    .split('\n')
                                    .map((l) => measureTextWidth(l, previewFont)),
                                )
                                if (typeof style.width === 'number' && widest > style.width) {
                                  style.width = widest + 2
                                }
                                if (te.input.align) {
                                  // The preview is a flex container and its text is one
                                  // shrink-to-fit anonymous item: textAlign only aligns
                                  // lines within that item, justifyContent moves the item
                                  // itself off main-start
                                  style.textAlign = te.input.align
                                  if (te.input.align === 'center') style.justifyContent = 'center'
                                  if (te.input.align === 'right') style.justifyContent = 'flex-end'
                                }
                                return (
                                  <Fragment key={te.id}>
                                    {te.cover && (
                                      <div
                                        className="pdf-textedit-cover"
                                        style={inflateCss(
                                          pdfRectToCss(
                                            geom,
                                            unionCover(te.input.rect, te.cover),
                                            scale,
                                          ),
                                          1.5,
                                        )}
                                      />
                                    )}
                                    <div
                                      className={`pdf-textedit-preview${
                                        selected?.kind === 'textEdit' && selected.id === te.id
                                          ? ' pdf-textedit-selected'
                                          : ''
                                      }`}
                                      style={style}
                                      title={editTextMode ? t('editTextHint') : t('removeMarkup')}
                                      onClick={(e) => {
                                        e.stopPropagation()
                                        if (editTextMode && !readOnly) {
                                          draftSelectedRef.current = false
                                          // Block edits reopen as one logical paragraph (the
                                          // stored newText is the wrapped form); leading is
                                          // unscaled back to the original font size
                                          const blk =
                                            te.input.origin && te.input.lineLeading
                                              ? {
                                                  leftPt: te.input.origin[0],
                                                  firstBaseline: te.input.origin[1],
                                                  widthPt: te.input.rect[2] - te.input.rect[0],
                                                  lineHeight:
                                                    te.input.lineLeading *
                                                    (te.input.fontSize /
                                                      (te.input.newFontSize ?? te.input.fontSize)),
                                                  align: te.input.align ?? ('left' as const),
                                                }
                                              : undefined
                                          const value = blk
                                            ? (te.input.blockSource ??
                                              joinBlockLines(te.input.newText.split('\n')))
                                            : te.input.newText
                                          // Selection colors are stored against the
                                          // committed newText; carry them back onto the
                                          // draft's logical text
                                          const hexRuns = (te.input.colorRuns ?? []).map((r) => ({
                                            start: r.start,
                                            end: r.end,
                                            color: rgb255ToHex(r.color),
                                          }))
                                          const onNew = hexRuns.length
                                            ? runsToColors(te.input.newText.length, hexRuns)
                                            : undefined
                                          setTextDraft({
                                            origIdx,
                                            rect: te.input.rect,
                                            oldText: te.input.oldText,
                                            fontSize: te.input.fontSize,
                                            value,
                                            charColors: onNew
                                              ? value === te.input.newText
                                                ? onNew
                                                : mapCharColors(te.input.newText, onNew, value)
                                              : undefined,
                                            size: te.input.newFontSize,
                                            color: te.input.newColor
                                              ? rgb255ToHex(te.input.newColor)
                                              : undefined,
                                            font: te.input.newFont,
                                            bold: te.input.newBold ? true : undefined,
                                            italic: te.input.newItalic ? true : undefined,
                                            editId: te.id,
                                            cover: te.cover,
                                            block: blk,
                                          })
                                        } else {
                                          setSelected({
                                            kind: 'textEdit',
                                            id: te.id,
                                            ...popupPos(e.clientX, e.clientY),
                                          })
                                        }
                                      }}
                                    >
                                      {te.input.colorRuns?.length ? (
                                        // One wrapper span = one flex item: the preview is a
                                        // row flex container, and bare segments would become
                                        // separate items laid out horizontally, breaking '\n'
                                        // stacking in multi-line previews
                                        <span>
                                          {colorSegments(
                                            te.input.newText,
                                            runsToColors(
                                              te.input.newText.length,
                                              te.input.colorRuns.map((r) => ({
                                                start: r.start,
                                                end: r.end,
                                                color: rgb255ToHex(r.color),
                                              })),
                                            ),
                                          ).map((seg, i) =>
                                            seg.color ? (
                                              <span key={i} style={{ color: seg.color }}>
                                                {seg.text}
                                              </span>
                                            ) : (
                                              <Fragment key={i}>{seg.text}</Fragment>
                                            ),
                                          )}
                                        </span>
                                      ) : (
                                        te.input.newText
                                      )}
                                    </div>
                                  </Fragment>
                                )
                              })}
                            {textDraft &&
                              textDraft.origIdx === origIdx &&
                              (() => {
                                const box = pdfRectToCss(geom, textDraft.rect, scale)
                                const fs = (textDraft.size ?? textDraft.fontSize) * scale * 0.92
                                const lines = textDraft.value.split('\n')
                                const draftCss = textDraft.font
                                  ? EDIT_FONT_BY_ID.get(textDraft.font)?.css
                                  : undefined
                                const bodyFamily = getComputedStyle(document.body).fontFamily
                                const blk = textDraft.block
                                const sizePt = textDraft.size ?? textDraft.fontSize
                                const draftStyle =
                                  `${textDraft.italic ? 'italic ' : ''}${textDraft.bold ? 'bold' : ''}`.trim()
                                // Block editor: width locks to the block so the textarea's
                                // soft wrap previews the reflow; height tracks the committed
                                // wrap count (in the block's own leading, plus headroom for
                                // the preview/commit measurement gap)
                                const leadPx = blk
                                  ? blk.lineHeight * (sizePt / textDraft.fontSize) * scale
                                  : fs * 1.2
                                const wrapCount = blk
                                  ? lines.reduce(
                                      (n, p) =>
                                        n +
                                        (p.trim()
                                          ? wrapText(
                                              p,
                                              blk.widthPt,
                                              sizePt,
                                              draftCss ?? bodyFamily,
                                              draftStyle,
                                            ).length
                                          : 1),
                                      0,
                                    )
                                  : lines.length
                                // Line editor grows with the longest line (measured in the
                                // editor's own font) so typed text stays visible; cap at the
                                // page's right edge, beyond which the textarea scrolls
                                const editorFont =
                                  `${draftStyle} ${fs}px ${draftCss ?? bodyFamily}`.trim()
                                // Selection-level colors: the textarea can't render
                                // mixed colors, so its text goes transparent and a
                                // metric-identical mirror behind the caret shows them
                                const draftColors = textDraft.charColors?.some((c) => c)
                                  ? textDraft.charColors
                                  : undefined
                                const longest = Math.max(
                                  ...lines.map((l) => measureTextWidth(l, editorFont)),
                                )
                                const pageEdgeCap = geomDispSize(geom).width * scale - box.left - 8
                                const editorWidth = blk
                                  ? box.width + 8
                                  : Math.min(
                                      Math.max(box.width, 120, longest + 12),
                                      Math.max(pageEdgeCap, box.width, 120),
                                    )
                                return (
                                  <>
                                    {textDraft.cover && (
                                      <div
                                        className="pdf-textedit-cover"
                                        style={inflateCss(
                                          pdfRectToCss(
                                            geom,
                                            unionCover(textDraft.rect, textDraft.cover),
                                            scale,
                                          ),
                                          1.5,
                                        )}
                                      />
                                    )}
                                    <div
                                      className="pdf-textedit-editor"
                                      style={{ left: box.left, top: box.top }}
                                      onClick={(e) => e.stopPropagation()}
                                      onBlur={(e) => {
                                        // Commit only when focus leaves the editor entirely —
                                        // clicking the style bar must not close the draft
                                        if (!e.currentTarget.contains(e.relatedTarget)) {
                                          commitTextDraft()
                                        }
                                      }}
                                    >
                                      <div className="pdf-textedit-bar">
                                        {editFonts.length > 0 && (
                                          <select
                                            className="pdf-textedit-fontsel"
                                            title={t('texteditFont')}
                                            value={textDraft.font ?? ''}
                                            onChange={(e) =>
                                              setTextDraft((d) =>
                                                d ? { ...d, font: e.target.value || undefined } : d,
                                              )
                                            }
                                          >
                                            <option value="">{t('texteditFontOriginal')}</option>
                                            {editFonts.map((id) => (
                                              <option key={id} value={id}>
                                                {EDIT_FONT_BY_ID.get(id)?.label ?? id}
                                              </option>
                                            ))}
                                          </select>
                                        )}
                                        <input
                                          className="pdf-textedit-sizenum"
                                          type="number"
                                          min={4}
                                          max={200}
                                          title={t('watermarkSize')}
                                          value={
                                            textDraft.size ??
                                            Math.round(textDraft.fontSize * 10) / 10
                                          }
                                          onChange={(e) => {
                                            const v = Number(e.target.value)
                                            if (v >= 1) {
                                              setTextDraft((d) => (d ? { ...d, size: v } : d))
                                            }
                                          }}
                                        />
                                        <button
                                          className={`pdf-textedit-toggle${textDraft.bold ? ' active' : ''}`}
                                          title={t('texteditBold')}
                                          onClick={() =>
                                            setTextDraft((d) =>
                                              d ? { ...d, bold: d.bold ? undefined : true } : d,
                                            )
                                          }
                                        >
                                          B
                                        </button>
                                        <button
                                          className={`pdf-textedit-toggle pdf-textedit-toggle-i${
                                            textDraft.italic ? ' active' : ''
                                          }`}
                                          title={t('texteditItalic')}
                                          onClick={() =>
                                            setTextDraft((d) =>
                                              d ? { ...d, italic: d.italic ? undefined : true } : d,
                                            )
                                          }
                                        >
                                          I
                                        </button>
                                        {DRAW_COLORS.map((c) => (
                                          <button
                                            key={c.name}
                                            className={`pdf-textedit-swatch${
                                              textDraft.color === rgbToHex(c.rgb) ? ' active' : ''
                                            }`}
                                            style={{ background: cssRgb(c.rgb) }}
                                            title={t('drawColor')}
                                            onClick={() => applyDraftColor(rgbToHex(c.rgb), true)}
                                          />
                                        ))}
                                        <label
                                          className="pdf-textedit-more"
                                          title={t('moreColors')}
                                        >
                                          <input
                                            type="color"
                                            value={textDraft.color ?? '#000000'}
                                            onChange={(e) => applyDraftColor(e.target.value, false)}
                                          />
                                        </label>
                                      </div>
                                      <textarea
                                        ref={draftTaRef}
                                        className={`pdf-textedit-input${blk ? ' pdf-textedit-block' : ''}`}
                                        style={{
                                          width: editorWidth,
                                          height: wrapCount * leadPx + (blk ? leadPx : 0) + 6,
                                          fontSize: fs,
                                          lineHeight: `${leadPx}px`,
                                          ...(blk && blk.align !== 'left'
                                            ? { textAlign: blk.align }
                                            : {}),
                                          // Document-content color (the user's pick), not chrome
                                          ...(textDraft.color ? { color: textDraft.color } : {}),
                                          ...(draftColors
                                            ? {
                                                color: 'transparent',
                                                caretColor:
                                                  textDraft.color ?? 'var(--pdf-textedit-ink)',
                                              }
                                            : {}),
                                          ...(draftCss ? { fontFamily: draftCss } : {}),
                                          ...(textDraft.bold ? { fontWeight: 700 } : {}),
                                          ...(textDraft.italic ? { fontStyle: 'italic' } : {}),
                                        }}
                                        value={textDraft.value}
                                        autoFocus
                                        onFocus={(e) => {
                                          if (!draftSelectedRef.current) {
                                            draftSelectedRef.current = true
                                            e.currentTarget.select()
                                          }
                                        }}
                                        onChange={(e) => {
                                          const v = e.target.value
                                          setTextDraft((d) =>
                                            d
                                              ? {
                                                  ...d,
                                                  value: v,
                                                  charColors: d.charColors?.some((c) => c)
                                                    ? spliceCharColors(d.value, d.charColors, v)
                                                    : undefined,
                                                }
                                              : d,
                                          )
                                        }}
                                        onScroll={(e) => {
                                          const g = draftGhostRef.current
                                          if (g) g.scrollLeft = e.currentTarget.scrollLeft
                                        }}
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
                                            e.preventDefault()
                                            commitTextDraft()
                                          } else if (e.key === 'Escape') {
                                            e.stopPropagation()
                                            setTextDraft(null)
                                          }
                                        }}
                                      />
                                      {draftColors && (
                                        <div
                                          ref={draftGhostRef}
                                          aria-hidden
                                          className={`pdf-textedit-ghost${
                                            blk ? ' pdf-textedit-block' : ''
                                          }`}
                                          style={{
                                            width: editorWidth,
                                            height: wrapCount * leadPx + (blk ? leadPx : 0) + 6,
                                            fontSize: fs,
                                            lineHeight: `${leadPx}px`,
                                            ...(blk && blk.align !== 'left'
                                              ? { textAlign: blk.align }
                                              : {}),
                                            color: textDraft.color ?? 'var(--pdf-textedit-ink)',
                                            ...(draftCss ? { fontFamily: draftCss } : {}),
                                            ...(textDraft.bold ? { fontWeight: 700 } : {}),
                                            ...(textDraft.italic ? { fontStyle: 'italic' } : {}),
                                          }}
                                        >
                                          {colorSegments(textDraft.value, draftColors).map(
                                            (seg, i) =>
                                              seg.color ? (
                                                <span key={i} style={{ color: seg.color }}>
                                                  {seg.text}
                                                </span>
                                              ) : (
                                                <Fragment key={i}>{seg.text}</Fragment>
                                              ),
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  </>
                                )
                              })()}
                            {(imageEdits.some((ie) => ie.input.pageIndex === origIdx) ||
                              (editImageMode &&
                                pageImages.some((ref) => ref.pageIndex === origIdx))) && (
                              <div
                                className={
                                  editTextMode || drawTool || pendingSign || imagePick
                                    ? 'pdf-imgedit-passive'
                                    : undefined
                                }
                              >
                                <ImageEditLayer
                                  geom={geom}
                                  scale={scale}
                                  edits={imageEdits.filter((ie) => ie.input.pageIndex === origIdx)}
                                  existing={
                                    editImageMode
                                      ? pageImages.filter(
                                          (ref) =>
                                            ref.pageIndex === origIdx &&
                                            !claimedImageKeys.has(
                                              `${ref.pageIndex}:${imageRectKey(ref.rect)}`,
                                            ),
                                        )
                                      : []
                                  }
                                  selectedId={selected?.kind === 'imageEdit' ? selected.id : null}
                                  selectedKey={
                                    selected?.kind === 'pageImage' &&
                                    selected.ref.pageIndex === origIdx
                                      ? imageRectKey(selected.ref.rect)
                                      : null
                                  }
                                  editHint={t('editImageHint')}
                                  onSelectEdit={(id, x, y) =>
                                    setSelected({ kind: 'imageEdit', id, ...popupPos(x, y) })
                                  }
                                  onSelectExisting={(ref, x, y) => {
                                    prefetchExistingPng(ref)
                                    setSelected({ kind: 'pageImage', ref, ...popupPos(x, y) })
                                  }}
                                  onRect={readOnly ? undefined : updateImageEditRect}
                                  onExistingRect={
                                    readOnly
                                      ? undefined
                                      : (ref, rect) => transformExisting(ref, rect)
                                  }
                                  existingPng={(ref) =>
                                    existingPngs.get(`${ref.pageIndex}:${imageRectKey(ref.rect)}`)
                                  }
                                  onExistingDragStart={prefetchExistingPng}
                                />
                              </div>
                            )}
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
                              onMove={readOnly ? undefined : moveDrawing}
                              onResize={readOnly ? undefined : resizeDrawing}
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
                      ? t('searchCount', {
                          current: searchCurClamped + 1,
                          total: activeMatches.length,
                        })
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
                <button
                  type="button"
                  title={t('highlight')}
                  onClick={() => applyMarkup('highlight')}
                >
                  <span className="sel-swatch sel-swatch-hl" />
                </button>
                <button
                  type="button"
                  title={t('underline')}
                  onClick={() => applyMarkup('underline')}
                >
                  <span className="sel-swatch sel-swatch-ul">U</span>
                </button>
                <button
                  type="button"
                  title={t('strikeout')}
                  onClick={() => applyMarkup('strikeout')}
                >
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
                {selectedImageLayer() !== null && (
                  <>
                    <button
                      type="button"
                      title={t('imageRotateCw')}
                      onClick={() => rotateSelected(1)}
                    >
                      <IconRotateCw />
                    </button>
                    <button
                      type="button"
                      title={t('imageRotateCcw')}
                      onClick={() => rotateSelected(-1)}
                    >
                      <IconRotateCcw />
                    </button>
                    {selected.kind === 'pageImage' && (
                      <button type="button" title={t('imageReplace')} onClick={startReplaceImage}>
                        <IconSwapImage />
                      </button>
                    )}
                    <span className="pdf-del-popup-sep" />
                    <button type="button" onClick={toggleImageLayer}>
                      {selectedImageLayer() === 'aboveText' ? <IconLayerDown /> : <IconLayerUp />}
                      {t(
                        selectedImageLayer() === 'aboveText'
                          ? 'imageLayerBelow'
                          : 'imageLayerAbove',
                      )}
                    </button>
                    <span className="pdf-del-popup-sep" />
                  </>
                )}
                <button type="button" className="pdf-del-popup-danger" onClick={deleteSelected}>
                  <IconTrash />
                  {t(
                    selected.kind === 'pageImage' || selected.kind === 'imageEdit'
                      ? 'deleteImage'
                      : 'deleteAnnotation',
                  )}
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
            {notice && (
              <div className="pdf-toast pdf-toast-notice">
                <span>{notice}</span>
                <button type="button" onClick={() => setNotice(null)}>
                  {t('ok')}
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

          <footer className="status-bar">
            <div className="status-left">
              <span className="status-item">
                {t('appPageOf', { current: currentPage, total: pageCount })}
              </span>
            </div>
            <div className="status-right">
              <button className="zoom-btn" title={t('zoomOut')} onClick={zoomOut}>
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={MIN_SCALE * 100}
                max={MAX_SCALE * 100}
                step={5}
                value={Math.round(scale * 100)}
                onChange={(e) => applyScale(Number(e.target.value) / 100, null)}
              />
              <button className="zoom-btn" title={t('zoomIn')} onClick={zoomIn}>
                +
              </button>
              <span className="zoom-value">{Math.round(scale * 100)}%</span>
            </div>
          </footer>
        </div>
      </div>
    </div>
  )
}
