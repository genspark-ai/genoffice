import { Fragment, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type {
  CSSProperties,
  MouseEvent as ReactMouseEvent,
  ReactElement,
  ReactNode,
  RefObject,
} from 'react'
// legacy build: the modern build relies on new APIs like Math.sumPrecise that the current
// Electron V8 lacks, making embedded font parsing fail and whole pages render as garbled raw char codes
import {
  AnnotationMode,
  GlobalWorkerOptions,
  TextLayer,
  getDocument,
} from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { AiPanel, GensparkMark } from './ai/AiPanel'
import type { PdfAiDeps } from './ai/tools'
import {
  MARKUP_COLORS,
  geomDispSize,
  pdfRectToCss,
  pdfToView,
  quadSetsMatch,
  quadToRect,
  selectionQuadsByPage,
  viewToPdf,
} from './annotations'
import type { LocalMarkup, PageGeom } from './annotations'
import { groupLineSpans } from './text-line'
import { DRAW_COLORS, DrawLayer, cssRgb } from './DrawLayer'
import { ColorPickerPopover } from './ColorPicker'
import type { DrawTool, LocalDrawing, SavedNotePin } from './DrawLayer'
import { NoteMarginColumn } from './NoteMargin'
import type { NoteMarginDraft, NoteMarginThread } from './NoteMargin'
import {
  buildNoteThreads,
  flattenThread,
  pendingNoteKey,
  threadSubtree,
  toSavedNote,
} from './note-threads'
import type { NoteInput, NoteThreadItem, PdfJsAnnotData, SavedNoteAnnot } from './note-threads'
import { FormLayer } from './FormLayer'
import {
  buildFormCatalog,
  documentFormFeatures,
  hasXfaMarker,
  visibleFormWidgets,
  type FormCatalog,
  type FormField,
  type FormWidget,
} from './form-catalog'
import { ImageEditLayer, imageRectKey } from './ImageEditLayer'
import type { LocalImageEdit } from './ImageEditLayer'
import { CropDialog, CutoutDialog } from './ImageDialogs'
import { cropRect, flipPixels, multiplyAlpha } from './image-bake'
import type { CropFractions } from './image-bake'
import type { PixelImage } from './cutout'
import { navAction } from './keyNav'
import { rowOfVisIdx, spreadRows, stepPage } from './spread'
import { LinkLayer } from './LinkLayer'
import { OutlinePanel } from './OutlinePanel'
import type { OutlineNode } from './OutlinePanel'
import { printPdf } from './print'
import { PropertiesDialog } from './PropertiesDialog'
import { SignatureDialog, fileToCanvas } from './SignatureDialog'
import type { SignatureData } from './SignatureDialog'
import { signatureDrawingForField } from './signature-field'
import {
  renderStaticFormMark,
  renderStaticFormText,
  type StaticFormFillKind,
} from './static-form-fill'
import { StampDialog } from './StampDialog'
import { buildStamps } from './stamps'
import type { HeaderFooterConfig, WatermarkConfig } from './stamps'
import { buildSearchIndex, searchInIndex } from './search'
import type { SearchIndex, SearchMatch } from './search'
import { mapDocFont, type DocFontStyle } from './doc-font'
import { groupPageBlocks, reflowOverflows, type TextBlock } from './text-block'
import {
  joinBlockLines,
  mapLineRangeToBlock,
  measurePt,
  spliceBlockText,
  unifyRadicals,
  wrapText,
} from './text-wrap'
import {
  colorRunsEqual,
  colorSegments,
  colorsToRuns,
  decodeStyle,
  encodeStyle,
  mapCharColors,
  patchStyle,
  runsToColors,
  spliceCharColors,
} from './color-runs'
import type { CharStyle } from './color-runs'
import { platformShortcuts } from '@genoffice/i18n'
import { useDismissablePopover } from '@genoffice/ui'
import { useI18n } from './i18n/locale'
import { useAutosave } from './useAutosave'
import { EDIT_FONTS } from '../shared/ipc'
import type {
  AnnotDeleteInput,
  DrawingInput,
  FormValueInput,
  ImageEditFailure,
  ImageEditInput,
  ImageLayer,
  MarkupType,
  MetadataInput,
  NoteEditInput,
  PageImageRef,
  StaticFormFillRecord,
  StampInput,
  TextEditFailure,
  TextEditInput,
  TextEditValidation,
  TextInsertFailure,
  TextInsertInput,
} from '../shared/ipc'

const EDIT_FONT_BY_ID = new Map<string, (typeof EDIT_FONTS)[number]>(
  EDIT_FONTS.map((f) => [f.id, f]),
)

/** pdf.js AnnotationType codes for the markup subtypes we can delete */
const MARKUP_TYPE_BY_ANNOT: Record<number, MarkupType> = {
  9: 'highlight',
  10: 'underline',
  12: 'strikeout',
}

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
  onRenderState,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  /** Unsaved rotation delta (clockwise degrees) */
  rotationDelta: number
  visible: boolean
  onRenderState: (doc: PDFDocumentProxy, pageNo: number, pending: boolean) => void
}) {
  const holderRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const holder = holderRef.current
    if (!holder) return
    // Offscreen pages still release their bitmap. For an in-place rerender (save
    // reload, zoom, rotation), keep the previous canvas visible until its replacement
    // is fully rendered so the page never flashes white between document instances.
    if (!visible) {
      holder.replaceChildren()
      // A page captured by the post-save barrier may scroll out before its replacement
      // renders. Its overlays are no longer mounted, so treat the cleared offscreen page
      // as settled instead of making the whole document wait for the timeout.
      onRenderState(doc, pageNo, false)
      return
    }
    // Visibility may change while a save reload is running. Register newly visible
    // pages dynamically so global preview cleanup cannot outrun their bitmap swap.
    onRenderState(doc, pageNo, true)
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
        // FormLayer renders interactive widgets as HTML. Exclude their saved appearance
        // streams from the page bitmap or filled values are drawn twice after a reload.
        annotationMode: AnnotationMode.ENABLE_FORMS,
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
      // Notify after the bitmap swap, but before the browser paints. A post-save
      // reload uses this to remove the matching edit previews in the same frame.
      onRenderState(doc, pageNo, false)
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
    }
  }, [doc, pageNo, scale, rotationDelta, visible, onRenderState])
  return <div ref={holderRef} className="pdf-page-content" />
}

/** Overlay for unsaved markups. Purely visual (pointer-events: none) so the text
 *  underneath stays selectable; clicking is handled by the page-level hit test. */
function MarkupOverlay({
  markups,
  geom,
  scale,
  selectedId,
}: {
  markups: LocalMarkup[]
  geom: PageGeom
  scale: number
  selectedId: string | null
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
const IconPreviousField = () => (
  <Icon>
    <rect x="6" y="4.5" width="12" height="15" rx="1.5" />
    <path d="M14.5 8 L10.5 12 L14.5 16" />
  </Icon>
)
const IconNextField = () => (
  <Icon>
    <rect x="6" y="4.5" width="12" height="15" rx="1.5" />
    <path d="M9.5 8 L13.5 12 L9.5 16" />
  </Icon>
)
const IconCompleteForm = () => (
  <Icon>
    <rect x="4.5" y="5" width="15" height="14" rx="1.5" />
    <path d="M8 12 L10.8 14.8 L16.5 9" />
  </Icon>
)
const IconFormText = () => (
  <Icon>
    <path d="M5 6 H19 M12 6 V19 M8.5 19 H15.5" />
  </Icon>
)
const IconFormCheck = () => (
  <Icon>
    <path d="M4.5 12.5 L9.5 17.5 L19.5 6.5" />
  </Icon>
)
const IconFormCross = () => (
  <Icon>
    <path d="M6 6 L18 18 M18 6 L6 18" />
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
const IconInsertBlank = () => (
  <Icon>
    <rect x="6.7" y="4.5" width="10.5" height="15" rx="1" strokeDasharray="2.2 1.8" />
    <path d="M11.95 9.5 V14.5 M9.45 12 H14.45" />
  </Icon>
)
const IconRotateAll = () => (
  <Icon>
    <rect x="8.5" y="8.5" width="7" height="9" rx="0.8" />
    <path d="M6.2 6.2 A 8.2 8.2 0 0 1 17.8 6.2" />
    <path d="M17.8 3.4 L17.8 6.4 L14.8 6.4" />
    <path d="M17.8 17.8 A 8.2 8.2 0 0 1 6.2 17.8" />
    <path d="M6.2 20.6 L6.2 17.6 L9.2 17.6" />
  </Icon>
)
const IconReverse = () => (
  <Icon>
    <path d="M9 5.5 L9 18 M6.5 8 L9 5.5 L11.5 8" />
    <path d="M15 6 L15 18.5 M12.5 16 L15 18.5 L17.5 16" />
  </Icon>
)
const IconSplitPdf = () => (
  <Icon>
    <rect x="4.5" y="6.5" width="6.2" height="11" rx="0.8" />
    <rect x="13.3" y="6.5" width="6.2" height="11" rx="0.8" />
    <path d="M12 4.6 V7 M12 9.4 V11.8 M12 14.2 V16.6 M12 19 V19.4" />
  </Icon>
)
const IconMergePdf = () => (
  <Icon>
    <path d="M4.5 6.5 H8.5 V17.5 H4.5 Z" />
    <path d="M15.5 6.5 H19.5 V17.5 H15.5 Z" />
    <path d="M10 9 L12 12 L10 15 M14 9 L12 12 L14 15" />
  </Icon>
)
const IconMergePages = () => (
  <Icon>
    <rect x="5" y="4.5" width="14" height="15" rx="1" />
    <path d="M12 4.5 V19.5 M5 12 H19" />
  </Icon>
)
/** Target paper sizes for "page size" (points, portrait) */
const PAPER_SIZES = [
  { label: 'A3', w: 842, h: 1191 },
  { label: 'A4', w: 595, h: 842 },
  { label: 'A5', w: 420, h: 595 },
  { label: 'Letter', w: 612, h: 792 },
  { label: 'Legal', w: 612, h: 1008 },
] as const

const IconReplacePages = () => (
  <Icon>
    <rect x="4.5" y="7.5" width="9" height="12" rx="1" />
    <rect x="10.5" y="4.5" width="9" height="12" rx="1" strokeDasharray="2.2 1.8" />
  </Icon>
)
const IconSplitPages = () => (
  <Icon>
    <rect x="5" y="4.5" width="14" height="15" rx="1" />
    <path d="M12 4.5 V19.5" strokeDasharray="2.2 1.8" />
    <path d="M8.2 12 L5.8 12 M6.6 10.8 L5.4 12 L6.6 13.2 M15.8 12 L18.2 12 M17.4 10.8 L18.6 12 L17.4 13.2" />
  </Icon>
)
const IconCropPages = () => (
  <Icon>
    <path d="M7.5 4.5 V16.5 H19.5" />
    <path d="M4.5 7.5 H16.5 V19.5" />
  </Icon>
)
const IconPageSize = () => (
  <Icon>
    <rect x="4.5" y="4.5" width="15" height="15" rx="1" />
    <rect x="8.5" y="8.5" width="7" height="9" rx="0.8" strokeDasharray="2 1.6" />
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
    <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
    <path d="M17 21v-8H7v8" />
    <path d="M7 3v5h8V3" />
  </IconRatio>
)

// ── selection-popup icons (14px; bring-forward / send-backward / trash) ──

const IconLayerUp = () => (
  <Icon size={18}>
    <rect x="9.5" y="4.5" width="10" height="10" rx="1" />
    <path d="M14.5 19.5 H5.5 A1 1 0 0 1 4.5 18.5 V9.5" />
  </Icon>
)
const IconLayerDown = () => (
  <Icon size={18}>
    <rect x="4.5" y="9.5" width="10" height="10" rx="1" />
    <path d="M9.5 4.5 H18.5 A1 1 0 0 1 19.5 5.5 V14.5" />
  </Icon>
)
const IconTrash = () => (
  <Icon size={18}>
    <path d="M4.5 6.5 H19.5" />
    <path d="M9 6.5 V5 A1 1 0 0 1 10 4 H14 A1 1 0 0 1 15 5 V6.5" />
    <path d="M6.5 6.5 L7.3 18.6 A1.4 1.4 0 0 0 8.7 19.9 H15.3 A1.4 1.4 0 0 0 16.7 18.6 L17.5 6.5" />
    <path d="M10.2 10 V16 M13.8 10 V16" />
  </Icon>
)
const IconRotateCw = () => (
  <Icon size={18}>
    <path d="M18.5 8.5 A7.5 7.5 0 1 0 19.5 12" />
    <path d="M19 4 V8.5 H14.5" />
  </Icon>
)
const IconRotateCcw = () => (
  <Icon size={18}>
    <path d="M5.5 8.5 A7.5 7.5 0 1 1 4.5 12" />
    <path d="M5 4 V8.5 H9.5" />
  </Icon>
)
const IconSwapImage = () => (
  <Icon size={18}>
    <rect x="4" y="6.5" width="11" height="11" rx="1" />
    <circle cx="7.5" cy="10" r="1.2" />
    <path d="M4.5 16 L8.5 12.5 L11 15 L12.5 13.5 L15 16" />
    <path d="M17 4.5 H19 A1 1 0 0 1 20 5.5 V13" />
    <path d="M18.2 11.2 L20 13 L21.8 11.2" />
  </Icon>
)
const IconFlipH = () => (
  <Icon size={18}>
    <path d="M12 3.5 V20.5" strokeDasharray="2.6 2.2" />
    <path d="M8.5 7 V17 L3.5 17 Z" />
    <path d="M15.5 7 V17 L20.5 17 Z" />
  </Icon>
)
const IconFlipV = () => (
  <Icon size={18}>
    <path d="M3.5 12 H20.5" strokeDasharray="2.6 2.2" />
    <path d="M7 8.5 H17 L17 3.5 Z" />
    <path d="M7 15.5 H17 L17 20.5 Z" />
  </Icon>
)
const IconCrop = () => (
  <Icon size={18}>
    <path d="M7 3.5 V17 H20.5" />
    <path d="M3.5 7 H17 V20.5" />
  </Icon>
)
const IconCutout = () => (
  <Icon size={18}>
    <path d="M13.5 6.5 L17.5 10.5 L8 20 H4 V16 Z" />
    <path d="M16 4 L20 8" />
    <path d="M18.5 12.5 L19.4 14.6 L21.5 15.5 L19.4 16.4 L18.5 18.5 L17.6 16.4 L15.5 15.5 L17.6 14.6 Z" />
  </Icon>
)
const IconOpacity = () => (
  <Icon size={18}>
    <path d="M12 3.5 C12 3.5 5.5 10 5.5 14.5 A6.5 6.5 0 0 0 18.5 14.5 C18.5 10 12 3.5 12 3.5 Z" />
    <path d="M12 18.2 A3.7 3.7 0 0 1 8.3 14.5" />
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

/** Committed IPC style runs → encoded-key runs over newText (the draft/preview form) */
const styleRunsToKeyRuns = (
  runs: NonNullable<TextEditInput['styleRuns']>,
): { start: number; end: number; color: string }[] =>
  runs.map((r) => ({
    start: r.start,
    end: r.end,
    color: encodeStyle({
      color: r.color ? rgb255ToHex(r.color) : undefined,
      font: r.font,
      size: r.size,
      bold: r.bold,
      italic: r.italic,
    }),
  }))

/** CSS of one styled segment in the editor mirror / pending preview. Explicit on/off
    overrides the inherited draft-level weight/slant; size scales like the host text. */
const styleSegCss = (s: CharStyle, scale: number): CSSProperties => ({
  ...(s.color ? { color: s.color } : {}),
  ...(s.font ? { fontFamily: EDIT_FONT_BY_ID.get(s.font)?.css } : {}),
  ...(s.size !== undefined ? { fontSize: s.size * scale * 0.92 } : {}),
  ...(s.bold !== undefined ? { fontWeight: s.bold ? 700 : 400 } : {}),
  ...(s.italic !== undefined ? { fontStyle: s.italic ? 'italic' : 'normal' } : {}),
})

/** Encoded-key runs → the IPC styleRuns the engine consumes */
const keyRunsToStyleRuns = (
  runs: { start: number; end: number; color: string }[],
): NonNullable<TextEditInput['styleRuns']> =>
  runs.map((r) => {
    const s = decodeStyle(r.color)
    return {
      start: r.start,
      end: r.end,
      color: s.color ? hexTo255(s.color) : undefined,
      font: s.font,
      size: s.size,
      bold: s.bold,
      italic: s.italic,
    }
  })

const rectsNear = (a: readonly number[], b: readonly number[], tolerance = 2): boolean =>
  a.length === 4 &&
  b.length === 4 &&
  a.every((value, index) => Math.abs(value - b[index]!) <= tolerance)

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
type RibbonTab = (typeof RIBBON_TABS)[number]['id'] | 'fillForm'

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

/** Width of the WPS-style comments margin beside the pages */
const NOTE_MARGIN_W = 300

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
  /** The run's base ink (display-only, from the draft's probe): the pending preview
      shows the document's real color when the edit doesn't change it */
  baseInk?: string
  /** Local look-alike of the run's own font (display-only, from the PostScript
      name): the pending preview reads like the document. Never sent to the engine. */
  baseFont?: DocFontStyle
  /** Accumulated block-move delta (PDF user space). Renderer metadata only: the preview
      and hover box draw at rect + moveBy while input.rect stays at the original position
      (it is the save-time match key). The engine-side position rides in input.translate
      (pure moves) or the shifted input.origin (moved rebuild edits). */
  moveBy?: [number, number]
}

/** Stable per-render key for a clustered block's rect (same idea as imageRectKey) */
const blockRectKey = (r: readonly number[]): string => r.map((v) => v.toFixed(2)).join(',')

const shiftRect = (
  r: readonly [number, number, number, number],
  d: readonly [number, number],
): [number, number, number, number] => [r[0] + d[0], r[1] + d[1], r[2] + d[0], r[3] + d[1]]

interface LocalTextInsert {
  id: string
  input: TextInsertInput
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

/** Pending text-insert preview style; shared by the canvas overlay and the thumbnail mirror */
const textInsertPreviewStyle = (
  insert: LocalTextInsert,
  geom: PageGeom,
  scale: number,
): CSSProperties => {
  const [vx, vy] = pdfToView(geom, insert.input.origin[0], insert.input.origin[1])
  const align = insert.input.align ?? 'left'
  const style: CSSProperties = {
    left: vx * scale,
    top: (vy - insert.input.fontSize) * scale,
    fontSize: insert.input.fontSize * scale * 0.92,
    lineHeight: insert.input.lineLeading ? `${insert.input.lineLeading * scale}px` : 1.2,
    color: `rgb(${insert.input.color.join(', ')})`,
    whiteSpace: 'pre',
    transform:
      align === 'center' ? 'translateX(-50%)' : align === 'right' ? 'translateX(-100%)' : undefined,
    textAlign: align,
  }
  if (insert.input.font) style.fontFamily = EDIT_FONT_BY_ID.get(insert.input.font)?.css
  if (insert.input.bold) style.fontWeight = 700
  if (insert.input.italic) style.fontStyle = 'italic'
  return style
}

/** Pending text-edit preview style + original-run cover; shared with the thumbnail mirror */
const textEditPreviewParts = (
  te: LocalTextEdit,
  geom: PageGeom,
  scale: number,
): { style: CSSProperties; coverStyle: CSSProperties | null } => {
  const fs = (te.input.newFontSize ?? te.input.fontSize) * scale * 0.92
  const lineCount = te.input.newText.split('\n').length
  const leadPx = te.input.lineLeading ? te.input.lineLeading * scale : fs * 1.2
  const style: CSSProperties = {
    // A moved block previews at its new position; input.rect stays at the original
    // (it is the save-time match key, and the cover below must hide the original ink)
    ...pdfRectToCss(geom, te.moveBy ? shiftRect(te.input.rect, te.moveBy) : te.input.rect, scale),
    fontSize: fs,
    ...(te.input.lineLeading ? { lineHeight: `${leadPx}px` } : {}),
  }
  if (te.input.newColor) {
    style.color = `rgb(${te.input.newColor.join(', ')})`
  } else if (te.baseInk) {
    style.color = te.baseInk
  }
  const baseF = te.input.newFont ? undefined : te.baseFont
  if (te.input.newFont) {
    style.fontFamily = EDIT_FONT_BY_ID.get(te.input.newFont)?.css
  } else if (baseF) {
    // Look-alike of the document's own face
    style.fontFamily = baseF.css
  }
  if (te.input.newBold) style.fontWeight = 700
  else if (baseF?.weight) style.fontWeight = baseF.weight
  if (te.input.newItalic) style.fontStyle = 'italic'
  else if (baseF?.italic) style.fontStyle = 'italic'
  if (lineCount > 1) {
    // Grow below the original rect, same leading the engine writes
    // (block edits carry the paragraph's own leading)
    style.height = lineCount * leadPx
    style.lineHeight = te.input.lineLeading ? `${leadPx}px` : 1.2
    style.alignItems = 'flex-start'
  } else if (typeof style.height === 'number' && fs + 2 > style.height) {
    // Enlarged single line: the engine keeps the run's baseline, so glyphs grow
    // upward past the original rect — extend the box up and bottom-align, or the
    // preview clips the taller glyphs (overflow: hidden)
    const grown = Math.ceil(fs) + 2
    style.top = (style.top as number) - (grown - style.height)
    style.height = grown
    style.alignItems = 'flex-end'
  }
  // The rebuilt run grows right past the original rect when the replacement is
  // longer; the preview must too, or the extra characters look cut off until
  // the save (overflow: hidden)
  const previewFont = `${te.input.newItalic || baseF?.italic ? 'italic ' : ''}${
    te.input.newBold ? 'bold ' : baseF?.weight ? `${baseF.weight} ` : ''
  }${fs}px ${
    (te.input.newFont && EDIT_FONT_BY_ID.get(te.input.newFont)?.css) ||
    baseF?.css ||
    getComputedStyle(document.body).fontFamily
  }`
  const widest = Math.max(
    ...te.input.newText.split('\n').map((l) => measureTextWidth(l, previewFont)),
  )
  if (typeof style.width === 'number' && widest > style.width) {
    style.width = widest + 2
  }
  if (te.input.align) {
    // The preview is a flex container and its text is one shrink-to-fit anonymous
    // item: textAlign only aligns lines within that item, justifyContent moves
    // the item itself off main-start
    style.textAlign = te.input.align
    if (te.input.align === 'center') style.justifyContent = 'center'
    if (te.input.align === 'right') style.justifyContent = 'flex-end'
  }
  const coverStyle = te.cover
    ? inflateCss(pdfRectToCss(geom, unionCover(te.input.rect, te.cover), scale), 1.5)
    : null
  return { style, coverStyle }
}

/** Styled-run children of a text-edit preview; shared with the thumbnail mirror */
const textEditPreviewContent = (te: LocalTextEdit, scale: number): ReactNode =>
  (te.input.styleRuns ?? te.input.colorRuns)?.length ? (
    // One wrapper span = one flex item: the preview is a row flex container, and
    // bare segments would become separate items laid out horizontally, breaking
    // '\n' stacking in multi-line previews
    <span>
      {colorSegments(
        te.input.newText,
        runsToColors(
          te.input.newText.length,
          styleRunsToKeyRuns(te.input.styleRuns ?? te.input.colorRuns ?? []),
        ),
      ).map((seg, i) => {
        if (!seg.color) return <Fragment key={i}>{seg.text}</Fragment>
        const s = decodeStyle(seg.color)
        return (
          <span key={i} style={styleSegCss(s, scale)}>
            {seg.text}
          </span>
        )
      })}
    </span>
  ) : (
    te.input.newText
  )

/** Read-only mirror of the pending-edit overlays, scaled into a page thumbnail so the
 *  sidebar tracks unsaved edits like the canvas does. Children render at scale 1 (display
 *  coords) and the container scales down; CSS kills all inner pointer events so thumbnail
 *  click/drag/context-menu behavior is untouched. */
function ThumbPendingOverlay({
  geom,
  k,
  markups,
  drawings,
  textEdits,
  textInserts,
  imageEdits,
  stamps,
}: {
  geom: PageGeom
  k: number
  markups: LocalMarkup[]
  drawings: LocalDrawing[]
  textEdits: LocalTextEdit[]
  textInserts: LocalTextInsert[]
  imageEdits: LocalImageEdit[]
  stamps: StampInput[]
}): ReactElement {
  const disp = geomDispSize(geom)
  const noop = () => {}
  return (
    <div
      className="pdf-thumb-overlay"
      style={{ width: disp.width, height: disp.height, transform: `scale(${k})` }}
    >
      {imageEdits.length > 0 && (
        <ImageEditLayer
          geom={geom}
          scale={1}
          edits={imageEdits}
          existing={[]}
          selectedId={null}
          selectedKey={null}
          editHint=""
          onSelectEdit={noop}
          onSelectExisting={noop}
        />
      )}
      {stamps.map((s, si) => (
        <img
          key={si}
          className="pdf-stamp-preview"
          src={`data:image/png;base64,${s.image}`}
          alt=""
          style={{ ...pdfRectToCss(geom, s.rect, 1), opacity: s.opacity ?? 1 }}
        />
      ))}
      <MarkupOverlay markups={markups} geom={geom} scale={1} selectedId={null} />
      {drawings.length > 0 && (
        <DrawLayer
          geom={geom}
          scale={1}
          pageWidth={disp.width}
          pageHeight={disp.height}
          drawings={drawings}
          savedNotes={[]}
          activeNoteKey={null}
          noteOpenTitle=""
          tool={null}
          color={[0, 0, 0]}
          strokeWidth={STROKE_WIDTH}
          selectedId={null}
          selectTitle=""
          onCommit={noop}
          onNoteAt={noop}
          onNoteOpen={noop}
          onSelect={noop}
        />
      )}
      {textInserts.map((insert) => (
        <div
          key={insert.id}
          className="pdf-textinsert-preview"
          style={textInsertPreviewStyle(insert, geom, 1)}
        >
          {insert.input.text}
        </div>
      ))}
      {textEdits.map((te) => {
        const { style, coverStyle } = textEditPreviewParts(te, geom, 1)
        return (
          <Fragment key={te.id}>
            {coverStyle && <div className="pdf-textedit-cover" style={coverStyle} />}
            <div className="pdf-textedit-preview" style={style}>
              {textEditPreviewContent(te, 1)}
            </div>
          </Fragment>
        )
      })}
    </div>
  )
}

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
  /** Selection-level styles, one encoded key per code unit of value (see encodeStyle);
      '' = base (the draft-level overrides ?? original). undefined/all-'' = uniform
      draft (the pre-existing whole-run behavior). */
  charStyles?: string[]
  /** Colors the document already draws the run with (async, from the open probe),
      pre-seeded into charStyles as color-only keys. A commit whose styles still equal
      these carries no *change* — they only ride along so a rebuild repaints them. */
  seedStyleRuns?: { start: number; end: number; color: string }[]
  /** The run's base ink in the document (async, from the open probe). Display-only:
      the editor/preview text shows the real color; never committed as a change. */
  seedInk?: string
  /** Local look-alike of the run's own font (async, from the dominant run's
      PostScript name). Display + reflow measurement only; never committed. */
  seedFont?: DocFontStyle
  /** EDIT_FONTS id; undefined = automatic rebuild font */
  font?: string
  /** Style toggles; true = on, undefined = off (resolved via font variants at save) */
  bold?: true
  italic?: true
  editId?: string
  /** Further pending edits folded into this block draft (besides editId); the
      commit replaces editId and removes these — they would overlap the block
      edit at save otherwise */
  foldedIds?: string[]
  /** The value the editor opened with when pending edits were folded in: an
      unmodified commit must keep those edits instead of converting them */
  foldBase?: string
  /** charStyles the fold seeded (styles of the folded line edits); an unmodified
      commit compares against these, not against empty */
  foldStyles?: string[]
  /** Ink bounds of the run being edited (async, from a dry-run validate) */
  cover?: [number, number, number, number]
  /** Present for paragraph (block) edits: the geometry the commit reflows into.
      lineHeight is the block's original leading at the original font size;
      bottomPt is the block's bottom edge in firstBaseline's (possibly moved)
      frame — the overflow guard measures growth against it. */
  block?: {
    leftPt: number
    firstBaseline: number
    widthPt: number
    lineHeight: number
    align: 'left' | 'center' | 'right'
    bottomPt: number
  }
  /** Carried from a reopened moved edit: the floating editor draws at rect + moveBy
      (where the preview sits) while rect itself stays the save-time match key */
  moveBy?: [number, number]
}

/** Seed a fresh draft's colors from the open probe's report of what the document
    already draws: the run's base ink (display-only) plus earlier saved selection
    colors. Selection colors only while the draft is pristine — the runs are offsets
    into oldText, and once typing starts they no longer align (the engine-side
    rebuild still preserves the colors on save). */
const seedDraftColors = (d: TextDraft, v: TextEditValidation): TextDraft => {
  let next = d
  // Near-white ink would vanish on the editor's white background; keep default ink
  if (v.baseColor && !next.color && !next.seedInk) {
    const [r, g, b] = v.baseColor
    if ((0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 <= 0.85)
      next = { ...next, seedInk: rgb255ToHex(v.baseColor) }
  }
  if (!v.colorRuns || v.colorRuns.length === 0) return next
  if (next.charStyles || next.seedStyleRuns || next.value !== next.oldText) return next
  const keyRuns = v.colorRuns.map((r) => ({
    start: r.start,
    end: r.end,
    color: encodeStyle({ color: rgb255ToHex(r.color) }),
  }))
  return {
    ...next,
    charStyles: runsToColors(next.oldText.length, keyRuns),
    seedStyleRuns: keyRuns,
  }
}

/** A markup annotation already saved in the file (read via pdf.js getAnnotations) */
interface SavedMarkupAnnot {
  pageIndex: number
  /** PDF object number (pdf.js id "123R" → 123) */
  objNum: number
  type: MarkupType
  quads: number[][]
  rect: [number, number, number, number]
}

/** Pending deletion of a saved markup or note annotation */
interface LocalAnnotDelete {
  id: string
  annot: SavedMarkupAnnot | SavedNoteAnnot
}

/** Pending in-place content edit of a note comment saved in the file.
    `annot.contents` stays the on-disk text — the save path matches by it. */
interface LocalNoteEdit {
  id: string
  annot: SavedNoteAnnot
  contents: string
}

interface EditSnapshot {
  markups: LocalMarkup[]
  annotDeletes: LocalAnnotDelete[]
  noteEdits: LocalNoteEdit[]
  drawings: LocalDrawing[]
  textEdits: LocalTextEdit[]
  textInserts: LocalTextInsert[]
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
  annotDeleteIds: Set<string>
  noteEditIds: Set<string>
  /** objNum → /Contents this save wrote via noteEdits. A re-edit of the same note made
      while the write was in flight still carries the pre-save on-disk text as its match
      key; the post-save subtraction retargets it to what the file now says. */
  noteEditWritten: Map<number, string>
  drawingIds: Set<string>
  textEditIds: Set<string>
  textInsertIds: Set<string>
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
  | {
      kind: 'markup' | 'drawing' | 'textEdit' | 'textInsert' | 'imageEdit'
      id: string
      x: number
      y: number
    }
  | { kind: 'savedMarkup'; annot: SavedMarkupAnnot; x: number; y: number }
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
const staticFormFillPlaceK = (): number => 1

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
      data-tip={title}
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
  const { lang, t } = useI18n()
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
  /** Pending deletions of markup annotations already saved in the file */
  const [annotDeletes, setAnnotDeletes] = useState<LocalAnnotDelete[]>([])
  /** Pending content rewrites of saved note comments (one entry per note, latest text) */
  const [noteEdits, setNoteEdits] = useState<LocalNoteEdit[]>([])
  /** Saved markup annotations per original page index, loaded lazily for visible pages (keyed to `doc`) */
  const [savedMarkups, setSavedMarkups] = useState<Map<number, SavedMarkupAnnot[]>>(new Map())
  /** Saved note (Text) comments per original page index, loaded in the same pass */
  const [savedNotes, setSavedNotes] = useState<Map<number, SavedNoteAnnot[]>>(new Map())
  /** Active comment thread: its margin card is expanded and linked to its pin */
  const [activeNote, setActiveNote] = useState<{ origIdx: number; rootKey: string } | null>(null)
  /** OS account name; the default author of new note comments */
  const [noteAuthor, setNoteAuthor] = useState('')
  useEffect(() => {
    window.pdfApi.getUsername().then(setNoteAuthor, () => {})
  }, [])
  const [highlightColor, setHighlightColor] = useState<[number, number, number]>(
    MARKUP_COLORS.highlight,
  )
  const [highlightColorOpen, setHighlightColorOpen] = useState(false)
  const [drawings, setDrawings] = useState<LocalDrawing[]>([])
  const [drawTool, setDrawTool] = useState<DrawTool | null>(null)
  const [textEdits, setTextEdits] = useState<LocalTextEdit[]>([])
  const [textInserts, setTextInserts] = useState<LocalTextInsert[]>([])
  const [pendingTextInsert, setPendingTextInsert] = useState<Omit<
    TextInsertInput,
    'pageIndex' | 'origin'
  > | null>(null)
  const [textInsertPointer, setTextInsertPointer] = useState<{
    pageIndex: number
    x: number
    y: number
  } | null>(null)
  useEffect(() => {
    if (!pendingTextInsert) setTextInsertPointer(null)
  }, [pendingTextInsert])
  const [editTextMode, setEditTextMode] = useState(false)
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null)
  /** Current draft for async callbacks (block-probe fallback runs after renders) */
  const textDraftRef = useRef<TextDraft | null>(null)
  textDraftRef.current = textDraft
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
  /** Border-drag of a clustered block (WPS-style move); client-space endpoints,
      converted to a PDF-space delta on release like the image-edit drag */
  const [blockDrag, setBlockDrag] = useState<{
    origIdx: number
    idx: number
    from: [number, number]
    to: [number, number]
  } | null>(null)
  /** When a grip drag releases, the browser still synthesizes a click at the drop
      point — which would open an editor for whatever text sits under it. Timestamped
      so a missed click (release outside the window) can't swallow a later real one.
      Shared with the pending-insert drag, whose stray click would pop the selection. */
  const blockDragReleaseAt = useRef(0)
  /** Drag-to-move of a pending inserted text (client-space endpoints, like blockDrag) */
  const [insertDrag, setInsertDrag] = useState<{
    id: string
    from: [number, number]
    to: [number, number]
  } | null>(null)
  const blockHoverRef = useRef<{ origIdx: number; idx: number } | null>(null)
  const clearBlockHover = () => {
    if (blockHoverRef.current) {
      blockHoverRef.current = null
      setBlockHover(null)
    }
  }
  /** Pending edit addressing this clustered block: block edits and block moves keep
      the block's own rect as their input.rect (the save-time match key), so those
      match by exact rect key. A LINE edit whose text covers the whole block (the
      single-line block case — its rect is the line span, not the block rect) owns
      the block just the same: moving/reopening the block must address that edit,
      not draft a colliding second one over the same objects. */
  const pendingEditFor = (origIdx: number, block: TextBlock): LocalTextEdit | undefined => {
    const pageEdits = textEdits.filter((te) => te.input.pageIndex === origIdx)
    if (pageEdits.length === 0) return undefined
    const key = blockRectKey(block.rect)
    const exact = pageEdits.find((te) => blockRectKey(te.input.rect) === key)
    if (exact) return exact
    const blockText = joinBlockLines(block.lines.map((l) => l.text)).replace(/\s+/g, '')
    if (!blockText) return undefined
    return pageEdits.find((te) => {
      const r = te.input.rect
      const cx = (r[0] + r[2]) / 2
      const cy = (r[1] + r[3]) / 2
      return (
        cx >= block.rect[0] &&
        cx <= block.rect[2] &&
        cy >= block.rect[1] &&
        cy <= block.rect[3] &&
        te.input.oldText.replace(/\s+/g, '') === blockText
      )
    })
  }

  /** Overlap of two PDF rects as a fraction of the smaller one (0 = disjoint) */
  const overlapOfSmaller = (a: readonly number[], b: readonly number[]): number => {
    const ox = Math.min(a[2]!, b[2]!) - Math.max(a[0]!, b[0]!)
    const oy = Math.min(a[3]!, b[3]!) - Math.max(a[1]!, b[1]!)
    if (ox <= 0 || oy <= 0) return 0
    const aArea = (a[2]! - a[0]!) * (a[3]! - a[1]!)
    const bArea = (b[2]! - b[0]!) * (b[3]! - b[1]!)
    return (ox * oy) / Math.min(aArea, bArea)
  }

  /** Pending edits that geometrically claim (part of) this block: rect overlap of at
      least half the smaller box. Catches everything the text matchers miss — a DOM
      visual line can join runs the clustering buckets into a neighboring block
      (slightly raised labels, superscripts), so a line edit's oldText covers MORE
      than the block and no text comparison lines up; likewise line edits inside a
      multi-line paragraph claim their row. Drag visuals let these previews follow
      the pointer and the ghost skips the rows they cover. */
  const blockClaimEdits = (origIdx: number, block: TextBlock): LocalTextEdit[] =>
    textEdits.filter(
      (te) => te.input.pageIndex === origIdx && overlapOfSmaller(te.input.rect, block.rect) >= 0.5,
    )

  const blockOverlapEdit = (origIdx: number, block: TextBlock): LocalTextEdit | undefined =>
    blockClaimEdits(origIdx, block)[0]

  /** The edit whose move/preview the block's drag visuals must follow. Single-line
      blocks fall back to geometric ownership (nothing to fold there anyway);
      multi-line blocks stay on the text matchers so paragraph folding keeps
      handling embedded line edits. */
  const blockOwnerEdit = (origIdx: number, b: TextBlock): LocalTextEdit | undefined =>
    pendingEditFor(origIdx, b) ?? (b.lines.length === 1 ? blockOverlapEdit(origIdx, b) : undefined)

  /** Where the block currently draws: its rect shifted by any pending move */
  const blockDrawRect = (origIdx: number, b: TextBlock): [number, number, number, number] => {
    const mv = blockOwnerEdit(origIdx, b)?.moveBy
    return mv ? shiftRect(b.rect, mv) : b.rect
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
      // Slop covers the border grips, which sit just OUTSIDE the block rect —
      // an exact test would clear the hover (and unmount the grip under the
      // pointer) the moment the pointer reaches them
      const pad = 8 / scale
      const hit = (b: TextBlock, p: number): boolean => {
        const r = blockDrawRect(origIdx, b)
        return px >= r[0] - p && px <= r[2] + p && py >= r[1] - p && py <= r[3] + p
      }
      // Sticky: keep the hovered block while the pointer stays in its padded
      // rect — two vertically-close blocks both claim the strip between them,
      // and without hysteresis the frame flips to the neighbour the moment the
      // pointer reaches a grip, so the press grabs the wrong block
      const held = cur?.origIdx === origIdx ? blocks[cur.idx] : undefined
      if (held && hit(held, pad)) return
      // Fresh acquisition: strictly-inside beats slop-only
      for (let i = 0; i < blocks.length && !next; i++) {
        if (hit(blocks[i]!, 0)) next = { origIdx, idx: i }
      }
      for (let i = 0; i < blocks.length && !next; i++) {
        if (hit(blocks[i]!, pad)) next = { origIdx, idx: i }
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
  /** Baseline metadata loaded from the PDF; pending imageEdits carry any changes. */
  const [savedStaticFormFills, setSavedStaticFormFills] = useState<StaticFormFillRecord[]>([])
  /** Picked image awaiting click-to-place (same overlay flow as signatures) */
  const [imagePick, setImagePick] = useState<Extract<SignatureData, { kind: 'image' }> | null>(null)
  const [pendingStaticFill, setPendingStaticFill] = useState<StaticFormFillKind | null>(null)
  const [staticTextDialog, setStaticTextDialog] = useState(false)
  const [staticTextPurpose, setStaticTextPurpose] = useState<'form' | 'insert'>('form')
  const [textInsertEditId, setTextInsertEditId] = useState<string | null>(null)
  const [staticText, setStaticText] = useState('')
  const [staticTextSize, setStaticTextSize] = useState(14)
  const [staticTextColor, setStaticTextColor] = useState('#111111')
  const [staticTextColorOpen, setStaticTextColorOpen] = useState(false)
  const [staticTextAlign, setStaticTextAlign] = useState<'left' | 'center' | 'right'>('left')
  useEffect(() => {
    if (!staticTextDialog) setStaticTextColorOpen(false)
  }, [staticTextDialog])
  const [staticTextEditTarget, setStaticTextEditTarget] = useState<
    | { kind: 'saved'; ref: PageImageRef; record: StaticFormFillRecord }
    | { kind: 'pending'; editId: string; record: StaticFormFillRecord }
    | null
  >(null)
  const imageFileRef = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (!imagePick) setPendingStaticFill(null)
  }, [imagePick])
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
  /** Initial caret/selection placement runs once per opened draft; refocusing after a
      style-bar click must keep the caret */
  const draftSelectedRef = useRef(false)
  /** Range to select in the next opened draft (WPS-style unified model: a click carries
      its caret as a collapsed range, a drag carries the dragged characters); null = the
      position is unknown — the caret goes to the end */
  const draftPreselectRef = useRef<[number, number] | null>(null)
  /** The open draft's textarea: style-bar color clicks read its selection (kept across blur) */
  const draftTaRef = useRef<HTMLTextAreaElement | null>(null)
  /** Shared color-picker popover on the text-edit style bar */
  const [draftColorOpen, setDraftColorOpen] = useState(false)
  /** Colored mirror behind the transparent-text textarea; scroll-synced to it */
  const draftGhostRef = useRef<HTMLDivElement | null>(null)
  const [drawColor, setDrawColor] = useState<[number, number, number]>(DRAW_COLORS[0]!.rgb)
  const [colorOpen, setColorOpen] = useState(false)
  /** Note just placed with the note tool; its content is typed into a margin draft card */
  const [noteDraft, setNoteDraft] = useState<{ origIdx: number; at: [number, number] } | null>(null)
  /** In-progress rewrite of an existing comment. Hoisted out of the margin card so a
      save can fold it in before the post-save reload tears the edit box down. */
  const [noteEditDraft, setNoteEditDraft] = useState<{
    origIdx: number
    rootKey: string
    itemKey: string
    text: string
  } | null>(null)
  // Deactivating the thread (close button, click elsewhere, another card) discards an
  // open comment edit — the same outcome as pressing Escape in the edit box
  useEffect(() => {
    if (noteEditDraft && activeNote?.rootKey !== noteEditDraft.rootKey) setNoteEditDraft(null)
  }, [activeNote, noteEditDraft])
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
  /** A /Sig widget selected from the form layer; confirmed signatures fit this rect directly. */
  const [signatureTarget, setSignatureTarget] = useState<FormWidget | null>(null)
  const [exporting, setExporting] = useState(false)
  const [formCatalog, setFormCatalog] = useState<FormCatalog | null>(null)
  const [formHasXfa, setFormHasXfa] = useState(false)
  const [documentEncrypted, setDocumentEncrypted] = useState(false)
  const [activeFormWidgetId, setActiveFormWidgetId] = useState<string | null>(null)
  const formControlRefs = useRef<Map<string, HTMLElement>>(new Map())
  const [formEdits, setFormEdits] = useState<Map<string, FormValueInput>>(new Map())
  const [rotations, setRotations] = useState<Map<number, number>>(new Map())
  const [deleted, setDeleted] = useState<Set<number>>(new Set())
  /** Markup bar over the current selection; quads (PDF space, keyed by original page
      index) drive the Word-style toggle state of the buttons */
  const [selPopup, setSelPopup] = useState<{
    x: number
    y: number
    quads: Map<number, number[][]>
  } | null>(null)
  const [selected, setSelected] = useState<AnnotSelection | null>(null)
  /** Transparency presets fold-out inside the image selection popup */
  const [opacityMenu, setOpacityMenu] = useState(false)
  useEffect(() => setOpacityMenu(false), [selected])
  const [deleteToast, setDeleteToast] = useState(false)
  const [deletedInsertedText, setDeletedInsertedText] = useState(false)
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
  const [splitDlg, setSplitDlg] = useState(false)
  const [splitInput, setSplitInput] = useState('1')
  const [splitInvalid, setSplitInvalid] = useState(false)
  const [mergePagesDlg, setMergePagesDlg] = useState(false)
  const [mergeCount, setMergeCount] = useState('2')
  const [mergeCountInvalid, setMergeCountInvalid] = useState(false)
  const [mergeDirection, setMergeDirection] = useState<'horizontal' | 'vertical'>('vertical')
  const [mergeSeparator, setMergeSeparator] = useState(false)
  const [replaceDlg, setReplaceDlg] = useState(false)
  const [replaceInput, setReplaceInput] = useState('')
  const [replaceInvalid, setReplaceInvalid] = useState(false)
  const [pageSizeDlg, setPageSizeDlg] = useState(false)
  const [splitPagesDlg, setSplitPagesDlg] = useState(false)
  /** Page-crop dialog: rendered page bitmap + which page it shows */
  const [pageCropDlg, setPageCropDlg] = useState<{ png: string; origIdx: number } | null>(null)
  const [cropAllPages, setCropAllPages] = useState(false)
  const coalesceKeyRef = useRef<string | null>(null)
  const passwordRef = useRef<string | undefined>(undefined)
  const fitModeRef = useRef<FitMode>('width')
  const scrollRef = useRef<HTMLDivElement>(null)
  const thumbsRef = useRef<HTMLDivElement>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  /** Local approximations of document fonts by pdf.js font id (see doc-font.ts);
      grows as edits/blocks touch pages, reset with the doc */
  const [docFonts, setDocFonts] = useState<Map<string, DocFontStyle>>(new Map())
  const docFontsRef = useRef(docFonts)
  docFontsRef.current = docFonts
  const searchIndexRef = useRef<{ doc: PDFDocumentProxy; promise: Promise<SearchIndex> } | null>(
    null,
  )
  const warnedXfaPathRef = useRef('')
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

  // Keep the current page's thumbnail in view as the main viewport drives currentPage
  useEffect(() => {
    if (sidebar !== 'thumbs') return
    const thumbEl = thumbsRef.current?.querySelector<HTMLElement>(`[data-idx="${currentPage - 1}"]`)
    thumbEl?.scrollIntoView({ block: 'nearest' })
  }, [currentPage, sidebar])

  // During a normal post-save reload, retain pending visual overlays until every
  // currently rendered page has swapped to the new document bitmap. Resolving the
  // barrier from PdfPage's render effect lets React clear those overlays before the
  // browser paints, avoiding the old-canvas-without-preview flash.
  const postSaveRenderWaitRef = useRef<{
    doc: PDFDocumentProxy
    pending: Set<number>
    finishScheduled: boolean
    finish: () => void
  } | null>(null)
  const pageRenderState = useCallback(
    (renderedDoc: PDFDocumentProxy, pageNo: number, pending: boolean) => {
      const wait = postSaveRenderWaitRef.current
      if (!wait || wait.doc !== renderedDoc) return
      if (pending) {
        wait.pending.add(pageNo)
        return
      }
      wait.pending.delete(pageNo)
      // Defer completion until all effects from the current visibility update have
      // reported. A page entering the viewport can then join before a departing page
      // releases the last item from the previous snapshot.
      if (wait.pending.size === 0 && !wait.finishScheduled) {
        wait.finishScheduled = true
        queueMicrotask(() => {
          wait.finishScheduled = false
          if (postSaveRenderWaitRef.current === wait && wait.pending.size === 0) wait.finish()
        })
      }
    },
    [],
  )

  const loadDoc = useCallback(
    async (
      path: string,
      previous: PDFDocumentProxy | null,
      saved?: SavedSnapshot,
      waitForPageNos: number[] = [],
    ) => {
      const data = await window.pdfApi.readFile(path)
      const bytes = new Uint8Array(data)
      setFormHasXfa(hasXfaMarker(bytes))
      if (!saved) {
        setFormCatalog(null)
        setSavedStaticFormFills([])
        setActiveFormWidgetId(null)
        formControlRefs.current.clear()
      }
      const loaded = await getDocument({
        data: bytes,
        password: passwordRef.current,
        ...DOC_OPTS,
      }).promise
      const metadata = await loaded.getMetadata()
      const documentInfo = metadata.info as {
        EncryptFilterName?: string | null
        IsXFAPresent?: boolean
      }
      const formFeatures = documentFormFeatures(documentInfo, bytes)
      setFormHasXfa(formFeatures.hasXfa)
      setDocumentEncrypted(formFeatures.encrypted)
      const all: PageSize[] = []
      const rots: number[] = []
      for (let i = 1; i <= loaded.numPages; i++) {
        const page = await loaded.getPage(i)
        // Unrotated size; display size is derived by geom from the total rotation
        const vp = page.getViewport({ scale: 1, rotation: 0 })
        all.push({ width: vp.width, height: vp.height })
        rots.push(page.rotate ?? 0)
      }
      try {
        setFormCatalog(await buildFormCatalog(loaded))
      } catch {
        setFormCatalog({ widgets: [], fields: new Map(), byPage: new Map() })
      }
      try {
        setSavedStaticFormFills(await window.pdfApi.listStaticFormFills(path))
      } catch {
        setSavedStaticFormFills([])
      }
      setSizes(all)
      setBaseRots(rots)
      let renderedPages: Promise<void> | null = null
      if (saved && waitForPageNos.length > 0) {
        postSaveRenderWaitRef.current?.finish()
        renderedPages = new Promise<void>((resolve) => {
          const wait = {
            doc: loaded,
            pending: new Set(waitForPageNos.filter((pageNo) => pageNo <= loaded.numPages)),
            finishScheduled: false,
            finish: () => {},
          }
          const timer = window.setTimeout(() => wait.finish(), 2000)
          wait.finish = () => {
            window.clearTimeout(timer)
            if (postSaveRenderWaitRef.current === wait) postSaveRenderWaitRef.current = null
            resolve()
          }
          postSaveRenderWaitRef.current = wait
          if (wait.pending.size === 0) wait.finish()
        })
      }
      setDoc(loaded)
      if (renderedPages) await renderedPages
      if (!saved) {
        setMarkups([])
        setAnnotDeletes([])
        setNoteEdits([])
        setDrawings([])
        setTextEdits([])
        setTextInserts([])
        setPendingTextInsert(null)
        setImageEdits([])
        setTextDraft(null)
        setStampCfg(null)
        setFormEdits(new Map())
        setSignatureTarget(null)
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
        setAnnotDeletes((prev) =>
          prev.flatMap((d) => {
            if (saved.annotDeleteIds.has(d.id)) return []
            const ni = remap.get(d.annot.pageIndex)
            if (ni === undefined) return []
            // The object number may be stale after the rewrite; the save path's
            // subtype+rect fallback still finds the annotation
            return [ni === d.annot.pageIndex ? d : { ...d, annot: { ...d.annot, pageIndex: ni } }]
          }),
        )
        setNoteEdits((prev) =>
          prev.flatMap((e) => {
            if (saved.noteEditIds.has(e.id)) return []
            const ni = remap.get(e.annot.pageIndex)
            if (ni === undefined) return []
            // An edit made while the write was in flight targets a note this save just
            // rewrote: match on the new on-disk contents, and drop the edit entirely if
            // it now says exactly what the file says
            const written = saved.noteEditWritten.get(e.annot.objNum)
            if (written !== undefined && e.contents === written) return []
            const contents = written ?? e.annot.contents
            if (ni === e.annot.pageIndex && contents === e.annot.contents) return [e]
            return [{ ...e, annot: { ...e.annot, pageIndex: ni, contents } }]
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
        setTextInserts((prev) =>
          prev.flatMap((insert) => {
            if (saved.textInsertIds.has(insert.id)) return []
            const ni = remap.get(insert.input.pageIndex)
            if (ni === undefined) return []
            return [
              ni === insert.input.pageIndex
                ? insert
                : { ...insert, input: { ...insert.input, pageIndex: ni } },
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

  /** pdf-lib cannot write encrypted files, including owner-protected files that open without a password. */
  const readOnly = status === 'ready' && (passwordRef.current !== undefined || documentEncrypted)

  useEffect(() => {
    if (
      activeFormWidgetId &&
      formCatalog &&
      !formCatalog.widgets.some((widget) => widget.id === activeFormWidgetId)
    ) {
      setActiveFormWidgetId(null)
    }
  }, [activeFormWidgetId, formCatalog])

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

  /** Cumulative row-top offset (with gaps), shared by scroll positioning and current-page calc */
  const rowTop = useCallback(
    (rowIdx: number) => {
      let y = PAGE_GAP
      for (let i = 0; i < rowIdx; i++) y += rowSize(rows[i]!).height * scale + PAGE_GAP
      return y
    },
    [rows, rowSize, scale],
  )

  /** Scroll offset that keeps the content point at the viewport top in place across a scale
   *  change. Row-exact: the fixed page gaps don't scale, so a plain scrollTop*ratio drifts
   *  by a full page once enough rows are above the viewport. Applied via layout effect so
   *  the write lands after React commits the resized pages (a rAF write raced the render
   *  and could get clamped against the old scrollHeight). */
  const pendingZoomScrollRef = useRef<number | null>(null)
  const anchorZoomScroll = useCallback(
    (nextScale: number) => {
      const el = scrollRef.current
      if (!el || scale <= 0 || rows.length === 0 || nextScale === scale) return
      const y = el.scrollTop
      let rowIdx = 0
      for (let i = 0; i < rows.length; i++) {
        if (rowTop(i) <= y) rowIdx = i
        else break
      }
      let top = PAGE_GAP
      for (let i = 0; i < rowIdx; i++) top += rowSize(rows[i]!).height * nextScale + PAGE_GAP
      // Only the page-content part of the within-row offset scales; anything outside the
      // page band (the leading margin above row 0, the inter-row gap) is fixed-size and
      // carries over unscaled — else zooming at the document top writes a non-zero scrollTop
      const within = y - rowTop(rowIdx)
      const content = Math.min(Math.max(within, 0), rowSize(rows[rowIdx]!).height * scale)
      pendingZoomScrollRef.current = top + (content / scale) * nextScale + (within - content)
    },
    [rows, rowSize, rowTop, scale],
  )
  useLayoutEffect(() => {
    if (pendingZoomScrollRef.current === null) return
    const el = scrollRef.current
    if (el) el.scrollTop = pendingZoomScrollRef.current
    pendingZoomScrollRef.current = null
  }, [scale])

  /** Reserve the WPS-style comments margin while the doc has notes or one is being added */
  const noteMarginOn = useMemo(
    () =>
      drawTool === 'note' ||
      noteDraft !== null ||
      drawings.some((d) => d.input.kind === 'note') ||
      (() => {
        const deletedObjs = new Set(annotDeletes.map((d) => d.annot.objNum))
        return [...savedNotes.values()].some((l) => l.some((a) => !deletedObjs.has(a.objNum)))
      })(),
    [drawTool, noteDraft, drawings, annotDeletes, savedNotes],
  )

  const recomputeFit = useCallback(() => {
    const mode = fitModeRef.current
    const el = scrollRef.current
    if (!mode || !el || rows.length === 0) return
    const dims = rows.map((r) => rowSize(r))
    const maxW = Math.max(...dims.map((s) => s.width))
    // The comments margin sits beside the pages and must fit inside the same viewport
    const availW = el.clientWidth - SCROLL_PAD * 2 - (noteMarginOn ? NOTE_MARGIN_W + PAGE_GAP : 0)
    const next =
      mode === 'width'
        ? clampScale(availW / maxW)
        : clampScale(
            Math.min(
              availW / maxW,
              (el.clientHeight - PAGE_GAP * 2) / Math.max(...dims.map((s) => s.height)),
            ),
          )
    anchorZoomScroll(next)
    setScale(next)
  }, [rows, rowSize, anchorZoomScroll, noteMarginOn])

  useEffect(() => {
    if (status !== 'ready') return
    recomputeFit()
    const el = scrollRef.current
    if (!el) return
    const ro = new ResizeObserver(recomputeFit)
    ro.observe(el)
    return () => ro.disconnect()
  }, [status, recomputeFit])

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

  const formWidgets = visibleFormWidgets(formCatalog, visList)
  const signedFormWidgetIds = useMemo(
    () =>
      new Set(
        drawings.flatMap((drawing) =>
          drawing.formWidgetId === undefined ? [] : [drawing.formWidgetId],
        ),
      ),
    [drawings],
  )
  const hasFillableForm = formWidgets.length > 0
  const activeFormIndex = formWidgets.findIndex((widget) => widget.id === activeFormWidgetId)

  const registerFormControl = (id: string, element: HTMLElement | null) => {
    if (element) formControlRefs.current.set(id, element)
    else formControlRefs.current.delete(id)
  }

  const focusFormWidget = (widget: FormWidget) => {
    const visibleIndex = visList.indexOf(widget.pageIndex)
    if (visibleIndex < 0) return
    setActiveFormWidgetId(widget.id)
    scrollToPage(visibleIndex + 1)
    setCurrentPage(visibleIndex + 1)
    setPageInput(String(visibleIndex + 1))
    let attempts = 0
    const focusWhenMounted = () => {
      const control = formControlRefs.current.get(widget.id)
      if (control) {
        control.focus()
        return
      }
      if (attempts++ < 8) requestAnimationFrame(focusWhenMounted)
    }
    requestAnimationFrame(focusWhenMounted)
  }

  const stepFormWidget = (direction: 1 | -1) => {
    if (formWidgets.length === 0) return
    const current = activeFormIndex >= 0 ? activeFormIndex : direction === 1 ? -1 : 0
    const next = (current + direction + formWidgets.length) % formWidgets.length
    focusFormWidget(formWidgets[next]!)
  }

  const formFieldFilled = (field: FormField): boolean => {
    const edit = formEdits.get(field.name)
    if (field.kind === 'checkbox') return edit ? !!edit.checked : field.checked
    // pdf.js deliberately hides the non-serializable /Sig value, and this app
    // does not verify certificate signatures. A visual Ink/Stamp signature must
    // therefore remain distinct from completion of a required digital-signature field.
    if (field.kind === 'signature') return false
    const value = edit && edit.kind !== 'checkbox' ? (edit.value ?? '') : field.value
    return value.trim().length > 0
  }

  const visibleFormFieldNames = new Set(formWidgets.map((widget) => widget.fieldName))
  const missingRequiredFields = [...(formCatalog?.fields.values() ?? [])].filter(
    (field) =>
      visibleFormFieldNames.has(field.name) &&
      field.required &&
      !field.readOnly &&
      !formFieldFilled(field),
  )

  /** Zoom keeping the content at the viewport top anchored (row-exact, see anchorZoomScroll) */
  const applyScale = (next: number, mode: FitMode) => {
    fitModeRef.current = mode
    const clamped = clampScale(next)
    anchorZoomScroll(clamped)
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

  const staticFormFills = useMemo(() => {
    const records = new Map(savedStaticFormFills.map((record) => [record.id, record]))
    for (const edit of imageEdits) {
      if (!edit.staticFill) continue
      if (edit.input.kind === 'deleteImage') {
        records.delete(edit.staticFill.id)
        continue
      }
      records.set(edit.staticFill.id, {
        ...edit.staticFill,
        pageIndex: edit.input.pageIndex,
        rect: edit.input.rect,
      })
    }
    return [...records.values()]
  }, [imageEdits, savedStaticFormFills])

  const dirty =
    markups.length > 0 ||
    annotDeletes.length > 0 ||
    noteEdits.length > 0 ||
    drawings.length > 0 ||
    textEdits.length > 0 ||
    textInserts.length > 0 ||
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
    annotDeletes,
    noteEdits,
    drawings,
    textEdits,
    textInserts,
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
    setAnnotDeletes(s.annotDeletes)
    setNoteEdits(s.noteEdits)
    setDrawings(s.drawings)
    setTextEdits(s.textEdits)
    setTextInserts(s.textInserts)
    setImageEdits(s.imageEdits)
    setTextDraft(null)
    // The comment being rewritten may not exist in the restored snapshot; confirming a
    // stale edit box would reapply text that undo just reverted
    setNoteEditDraft(null)
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
    setDocFonts(new Map())
    clearBlockHover()
    // The line affordance is just as stale after a reload (it hangs at the last
    // pre-save pointer position until the next mousemove)
    clearLineHover()
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
      const grouped = new Map<number, TextBlock[]>()
      for (const i of missing) {
        const entry = entries[i]
        grouped.set(i, entry ? groupPageBlocks(entry) : [])
      }
      setPageBlocks((prev) => {
        const next = new Map(prev)
        for (const [i, blocks] of grouped) if (!next.has(i)) next.set(i, blocks)
        return next
      })
      // Resolve the lines' document fonts up front so the border-drag ghost can
      // read like the page the moment a drag starts
      for (const [i, blocks] of grouped)
        for (const b of blocks) for (const l of b.lines) if (l.font) void resolveDocFont(i, l.font)
    })
    return () => {
      stale = true
    }
  }, [editTextMode, readOnly, doc, visibleRows, rows, pageBlocks, getSearchIndex])

  /** Saved markup/note annotations are keyed to the loaded doc; drop them on save-reload */
  useEffect(() => {
    setSavedMarkups(new Map())
    setSavedNotes(new Map())
    setActiveNote(null)
    setNoteDraft(null)
  }, [doc])

  /** Load saved markup + note annotations for pages scrolled into view: markups so
      clicking one can select it for deletion, notes so their comment threads show.
      Runs for read-only docs too (comments are viewable). Settles the same way as
      the paragraph-box effect. */
  useEffect(() => {
    if (!doc) return
    const missing: number[] = []
    for (const r of visibleRows)
      for (const i of rows[r] ?? []) if (!savedMarkups.has(i)) missing.push(i)
    if (missing.length === 0) return
    let stale = false
    void (async () => {
      const markupEntries: [number, SavedMarkupAnnot[]][] = []
      const noteEntries: [number, SavedNoteAnnot[]][] = []
      for (const origIdx of missing) {
        let markupList: SavedMarkupAnnot[] = []
        let noteList: SavedNoteAnnot[] = []
        try {
          const page = await doc.getPage(origIdx + 1)
          const annots = (await page.getAnnotations()) as (PdfJsAnnotData & {
            quadPoints?: Float32Array | null
          })[]
          markupList = annots.flatMap((a) => {
            const type = MARKUP_TYPE_BY_ANNOT[a.annotationType]
            // Only ref-backed annots can be addressed for deletion (id "123R" → object 123)
            const objNum = /^(\d+)R$/.exec(a.id)
            if (!type || !objNum || !a.quadPoints || a.quadPoints.length < 8) return []
            const quads: number[][] = []
            for (let q = 0; q + 8 <= a.quadPoints.length; q += 8)
              quads.push([...a.quadPoints.slice(q, q + 8)])
            return [
              {
                pageIndex: origIdx,
                objNum: Number(objNum[1]),
                type,
                quads,
                rect: [a.rect[0]!, a.rect[1]!, a.rect[2]!, a.rect[3]!] as [
                  number,
                  number,
                  number,
                  number,
                ],
              },
            ]
          })
          noteList = annots.flatMap((a) => {
            const note = toSavedNote(a, origIdx)
            return note ? [note] : []
          })
        } catch {
          /* page unreadable; no saved annotations to offer */
        }
        markupEntries.push([origIdx, markupList])
        noteEntries.push([origIdx, noteList])
      }
      if (!stale) {
        setSavedMarkups((prev) => new Map([...prev, ...markupEntries]))
        setSavedNotes((prev) => new Map([...prev, ...noteEntries]))
      }
    })()
    return () => {
      stale = true
    }
  }, [doc, visibleRows, rows, savedMarkups])

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

  /** Selection quads in PDF space keyed by original page index; null when nothing usable */
  const selectionQuads = (): Map<number, number[][]> | null => {
    const el = scrollRef.current
    if (!el) return null
    const byVisPage = selectionQuadsByPage(
      el,
      visList.map((i) => pageGeom(i)),
      scale,
    )
    if (!byVisPage) return null
    const quads = new Map<number, number[][]>()
    for (const [visIdx, q] of byVisPage) {
      const origIdx = visList[visIdx]
      if (origIdx !== undefined) quads.set(origIdx, q)
    }
    return quads.size > 0 ? quads : null
  }

  /** Mouse released over selected text → show the markup bar centered above the selection box (below if it doesn't fit) */
  const handleMouseUp = () => {
    // In edit-text mode a drag means "choose the characters to edit" (the click
    // after mouseup opens the editor preselected), not the markup popup
    if (editTextMode && !readOnly) return
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
      const quads = selectionQuads()
      if (!quads) return
      setSelPopup({
        x: Math.min(Math.max(box.left + box.width / 2, 70), window.innerWidth - 70),
        y: box.top >= 52 ? box.top - 44 : Math.min(box.bottom + 8, window.innerHeight - 44),
        quads,
      })
    }, 0)
  }

  /** The existing markup of `type` whose quads equal this page selection (tolerance-
      based): pending unsaved markups first, then annotations saved in the file
      (minus ones already pending deletion). Null = the selection isn't marked. */
  const markupMatching = (
    origIdx: number,
    type: MarkupType,
    quads: number[][],
  ): { pending: LocalMarkup } | { saved: SavedMarkupAnnot } | null => {
    const pending = markups.find(
      (m) => m.pageIndex === origIdx && m.type === type && quadSetsMatch(m.quads, quads),
    )
    if (pending) return { pending }
    const pendingDeleted = new Set(annotDeletes.map((d) => d.annot.objNum))
    const saved = (savedMarkups.get(origIdx) ?? []).find(
      (a) => a.type === type && !pendingDeleted.has(a.objNum) && quadSetsMatch(a.quads, quads),
    )
    return saved ? { saved } : null
  }

  /** Word-style toggle: apply `type` to the selection, or remove it when the whole
      selection already carries it. Selection and bar survive so more markup types can
      be stacked/toggled on the same text; clicking anywhere else dismisses them. */
  const applyMarkup = (type: MarkupType) => {
    if (readOnly) return
    const selQuads = selectionQuads()
    if (!selQuads) {
      setSelPopup(null)
      return
    }
    const matches = [...selQuads].map(
      ([origIdx, quads]) => [origIdx, quads, markupMatching(origIdx, type, quads)] as const,
    )
    if (matches.every(([, , m]) => m !== null)) {
      // Every page of the selection is already marked → the click removes
      pushUndo()
      const pendingIds = new Set(
        matches.flatMap(([, , m]) => (m && 'pending' in m ? [m.pending.id] : [])),
      )
      const savedHits = matches.flatMap(([, , m]) => (m && 'saved' in m ? [m.saved] : []))
      if (pendingIds.size > 0) setMarkups((prev) => prev.filter((m) => !pendingIds.has(m.id)))
      if (savedHits.length > 0)
        setAnnotDeletes((prev) => [...prev, ...savedHits.map((annot) => ({ id: newId(), annot }))])
      return
    }
    // Pages already carrying the markup are skipped, not duplicated (Word semantics:
    // applying to a partially-marked selection marks the rest)
    const added: LocalMarkup[] = matches.flatMap(([origIdx, quads, m]) =>
      m
        ? []
        : [
            {
              id: `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
              pageIndex: origIdx,
              type,
              color: type === 'highlight' ? highlightColor : MARKUP_COLORS[type],
              quads,
            },
          ],
    )
    if (added.length === 0) return
    pushUndo()
    setMarkups((prev) => [...prev, ...added])
  }

  /** Markup types the whole current selection already carries — shown as pressed
      buttons in the bar (clicking one removes the markup) */
  const activeMarkupTypes = useMemo(() => {
    const active = new Set<MarkupType>()
    if (!selPopup) return active
    for (const type of ['highlight', 'underline', 'strikeout'] as const) {
      if ([...selPopup.quads].every(([origIdx, quads]) => markupMatching(origIdx, type, quads)))
        active.add(type)
    }
    return active
    // markupMatching reads markups/annotDeletes/savedMarkups; they are all listed
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selPopup, markups, annotDeletes, savedMarkups])

  // ── Annotation selection: click selects, deletion is explicit (delete popup / Delete key) ──

  /** Clamp the delete popup anchor into the window, preferring a spot above the click (same rules as the markup bar) */
  const popupPos = (x: number, y: number) => ({
    x: Math.min(Math.max(x, 70), window.innerWidth - 70),
    y: y >= 96 ? y - 48 : Math.min(y + 12, window.innerHeight - 44),
  })

  /** Floating bars are anchored to content (a click point, an edited text run), so near a
   *  window edge their natural position can overflow — widths vary by content, so after
   *  render measure the real box and shift it back inside via margin-left */
  const clampFloatingBar = (el: HTMLElement | null) => {
    if (!el) return
    el.style.marginLeft = '0px'
    const r = el.getBoundingClientRect()
    const pad = 8
    const shift = r.left < pad ? pad - r.left : Math.min(0, window.innerWidth - pad - r.right)
    el.style.marginLeft = `${shift}px`
  }
  const delPopupRef = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    clampFloatingBar(delPopupRef.current)
  }, [selected])
  const textEditBarRef = useRef<HTMLDivElement>(null)
  /** These default to floating above their anchor; when the anchor sits at the top of the
   *  scroll viewport that position is clipped by the scroll container — flip below instead */
  const flipBelowIfClipped = (el: HTMLElement | null, offset: string) => {
    if (!el) return
    el.style.top = ''
    el.style.bottom = ''
    const clipTop = scrollRef.current?.getBoundingClientRect().top ?? 0
    if (el.getBoundingClientRect().top < clipTop) {
      el.style.bottom = 'auto'
      el.style.top = offset
    }
  }
  const placeTextEditBars = () => {
    flipBelowIfClipped(textEditBarRef.current, 'calc(100% + 5px)')
    clampFloatingBar(textEditBarRef.current)
  }
  // textDraft dep: the font select's width follows the chosen font, so re-clamp on change
  useLayoutEffect(placeTextEditBars, [textDraft])
  // The draft survives scrolling (unlike selections), so the anchor can move back under the
  // scrollport top after placement — re-run the flip/clamp as the viewport scrolls
  useEffect(() => {
    if (!textDraft) return
    const el = scrollRef.current
    if (!el) return
    el.addEventListener('scroll', placeTextEditBars, { passive: true })
    return () => el.removeEventListener('scroll', placeTextEditBars)
    // placeTextEditBars reads only refs; re-subscribing per draft is enough
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [textDraft])

  /** Markup overlays don't take pointer events (text under them must stay selectable),
   *  so a bare click on the page content hit-tests them here. Only a true click counts:
   *  a drag that produced a text selection goes to the markup bar instead. Pending
   *  (unsaved) markups win over saved annotations; last added wins within each group,
   *  matching visual stacking order. */
  const handlePageClick = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>) => {
    // Only clicks that land on the rendered page (canvas/text layer); overlays like
    // draw shapes, note pins, stamps and previews handle their own selection
    if (!(e.target as Element).closest?.('.pdf-page-content')) return
    const sel = window.getSelection()
    if (sel && !sel.isCollapsed) return
    const box = e.currentTarget.getBoundingClientRect()
    const g = pageGeom(origIdx)
    const [px, py] = viewToPdf(g, (e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
    const hitQuads = (quads: number[][]) =>
      quads.some((q) => {
        const r = quadToRect(q)
        return px >= r[0] && px <= r[2] && py >= r[1] && py <= r[3]
      })
    const select = (selection: AnnotSelection) => {
      e.stopPropagation() // keep the scroll container from clearing the selection we just set
      setSelected(selection)
    }
    const at = popupPos(e.clientX, e.clientY)
    const onPage = markups.filter((m) => m.pageIndex === origIdx)
    for (let i = onPage.length - 1; i >= 0; i--) {
      const m = onPage[i]!
      if (hitQuads(m.quads)) return select({ kind: 'markup', id: m.id, ...at })
    }
    if (readOnly) return
    // Markup annotations already saved in the file (skipping ones pending deletion)
    const pendingDeleted = new Set(annotDeletes.map((d) => d.annot.objNum))
    const saved = savedMarkups.get(origIdx) ?? []
    for (let i = saved.length - 1; i >= 0; i--) {
      const a = saved[i]!
      if (pendingDeleted.has(a.objNum) || !hitQuads(a.quads)) continue
      return select({ kind: 'savedMarkup', annot: a, ...at })
    }
  }

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

  /** Line edit that can fold into a paragraph draft: anything without a position of
      its own (block edits/moves carry origin/translate and are matched by rect key
      instead). Style overrides fold too — they become selection styles of the draft. */
  const isFoldableLineEdit = (i: TextEditInput) => i.origin === undefined && !i.translate

  /** A line edit's styling as per-code-unit encoded keys over its newText: the
      whole-edit overrides (newColor/newFont/…) as the base, with the selection
      styleRuns patched on top. undefined = the edit carries no styling. */
  const lineEditCharStyles = (i: TextEditInput): string[] | undefined => {
    const baseKey = encodeStyle({
      color: i.newColor ? rgb255ToHex(i.newColor) : undefined,
      font: i.newFont,
      size: i.newFontSize,
      bold: i.newBold ? true : undefined,
      italic: i.newItalic ? true : undefined,
    })
    const keyRuns = styleRunsToKeyRuns(i.styleRuns ?? i.colorRuns ?? [])
    if (!baseKey && keyRuns.length === 0) return undefined
    const perChar = runsToColors(i.newText.length, keyRuns)
    return perChar.map((k) => (k ? patchStyle(baseKey, decodeStyle(k)) : baseKey))
  }

  /** Pending line edits inside `block` folded into its paragraph text. Opening the
      paragraph over them with the original text would hide the user's changes, and
      committing would create a second edit over the same objects — skipped at save
      as overlapping. Styled line edits fold too: their styles come back as the
      draft's selection styles (charStyles over `value`). null = nothing to fold or
      some pending edit inside the block cannot fold (callers keep their previous
      behavior then). */
  const foldBlockValue = (
    origIdx: number,
    block: TextBlock,
    blockText: string,
  ): { value: string; editId: string; foldedIds: string[]; charStyles?: string[] } | null => {
    const inside = textEdits.filter((e) => {
      if (e.input.pageIndex !== origIdx) return false
      const r = e.input.rect
      const cx = (r[0] + r[2]) / 2
      const cy = (r[1] + r[3]) / 2
      return (
        cx >= block.rect[0] && cx <= block.rect[2] && cy >= block.rect[1] && cy <= block.rect[3]
      )
    })
    if (inside.length === 0) return null
    if (!inside.every((e) => isFoldableLineEdit(e.input))) return null
    // Non-space offset of each block line inside blockText, to disambiguate a
    // repeated oldText toward the line the edit actually sits on
    const offsets: number[] = []
    let acc = 0
    for (const l of block.lines) {
      offsets.push(acc)
      acc += l.text.replace(/\s+/g, '').length
    }
    const outRanges: [number, number][] = []
    const folded = spliceBlockText(
      blockText,
      inside.map((e) => {
        const cy = (e.input.rect[1] + e.input.rect[3]) / 2
        const li = block.lines.findIndex((l) => cy >= l.rect[1] && cy <= l.rect[3])
        return { oldText: e.input.oldText, newText: e.input.newText, hint: offsets[li] ?? 0 }
      }),
      outRanges,
    )
    if (folded === null) return null
    // Carry each folded edit's styles onto the paragraph value at its spliced range
    let styles: string[] | undefined
    for (const [k, e] of inside.entries()) {
      const perChar = lineEditCharStyles(e.input)
      if (!perChar) continue
      const [s0] = outRanges[k]!
      // The splice inserted unifyRadicals(newText); map the newText-space keys over
      // when folding changed the code units (rare radical-variant case)
      const unified = unifyRadicals(e.input.newText)
      const onUnified =
        unified === e.input.newText ? perChar : mapCharColors(e.input.newText, perChar, unified)
      styles ??= new Array<string>(folded.length).fill('')
      for (let i = 0; i < onUnified.length && s0 + i < folded.length; i++)
        styles[s0 + i] = onUnified[i]!
    }
    return {
      value: folded,
      editId: inside[0]!.id,
      foldedIds: inside.slice(1).map((e) => e.id),
      charStyles: styles?.some((k) => k) ? styles : undefined,
    }
  }

  /** Map a page font to its local look-alike via the PostScript name on the
      pdf.js font object (loaded once the page has rendered). Cached per doc;
      unresolved ids are retried on the next call. */
  const resolveDocFont = async (origIdx: number, fontId: string): Promise<DocFontStyle | null> => {
    const cached = docFontsRef.current.get(fontId)
    if (cached) return cached
    if (!doc) return null
    try {
      const page = await doc.getPage(origIdx + 1)
      const f = page.commonObjs.get(fontId) as { name?: string } | null
      if (!f?.name) return null
      const style = mapDocFont(f.name)
      setDocFonts((prev) => new Map(prev).set(fontId, style))
      return style
    } catch {
      // Font object not resolved yet (page not rendered) — retried next call
      return null
    }
  }

  /** Seed the draft with the document's own face: the font id covering the most
      characters inside the draft rect names it. Async like the ink probe; rect
      identity pins the draft. */
  const seedDraftFont = (origIdx: number, rect: [number, number, number, number]) => {
    const index = getSearchIndex()
    if (!index) return
    void index.then(async (entries) => {
      const items = entries[origIdx]?.items
      if (!items) return
      const weight = new Map<string, number>()
      for (const it of items) {
        if (!it.font) continue
        const cx = it.x + it.w / 2
        const cy = it.y + it.h / 2
        if (cx < rect[0] || cx > rect[2] || cy < rect[1] || cy > rect[3]) continue
        weight.set(it.font, (weight.get(it.font) ?? 0) + (it.end - it.start))
      }
      let best: string | undefined
      let bestN = 0
      for (const [f, n] of weight)
        if (n > bestN) {
          bestN = n
          best = f
        }
      if (!best) return
      const style = await resolveDocFont(origIdx, best)
      if (!style) return
      setTextDraft((d) =>
        d && d.origIdx === origIdx && d.rect === rect && !d.seedFont
          ? { ...d, seedFont: style }
          : d,
      )
    })
  }

  /** Rebuild the floating editor's draft from a pending edit (reopen). Shared by the
      preview click and by page clicks landing on a block that already carries a
      pending block edit or move — committing a *second* edit over the same objects
      would only be skipped as overlapping at save. */
  const pendingDraft = (te: LocalTextEdit): TextDraft => {
    // Block edits reopen as one logical paragraph (the stored newText is the wrapped
    // form); leading is unscaled back to the original font size
    const blk =
      te.input.origin && te.input.lineLeading
        ? {
            leftPt: te.input.origin[0],
            firstBaseline: te.input.origin[1],
            widthPt: te.input.rect[2] - te.input.rect[0],
            lineHeight:
              te.input.lineLeading *
              (te.input.fontSize / (te.input.newFontSize ?? te.input.fontSize)),
            align: te.input.align ?? ('left' as const),
            // rect is the unmoved match key; shift the bottom into the moved frame
            bottomPt: te.input.rect[1] + (te.moveBy?.[1] ?? 0),
          }
        : undefined
    const value = blk
      ? (te.input.blockSource ?? joinBlockLines(te.input.newText.split('\n')))
      : te.input.newText
    // Selection styles are stored against the committed newText; carry them back
    // onto the draft's logical text
    const keyRuns = styleRunsToKeyRuns(te.input.styleRuns ?? te.input.colorRuns ?? [])
    const onNew = keyRuns.length ? runsToColors(te.input.newText.length, keyRuns) : undefined
    return {
      origIdx: te.input.pageIndex,
      rect: te.input.rect,
      oldText: te.input.oldText,
      fontSize: te.input.fontSize,
      value,
      charStyles: onNew
        ? value === te.input.newText
          ? onNew
          : mapCharColors(te.input.newText, onNew, value)
        : undefined,
      size: te.input.newFontSize,
      color: te.input.newColor ? rgb255ToHex(te.input.newColor) : undefined,
      font: te.input.newFont,
      bold: te.input.newBold ? true : undefined,
      italic: te.input.newItalic ? true : undefined,
      editId: te.id,
      cover: te.cover,
      seedInk: te.baseInk,
      seedFont: te.baseFont,
      block: blk,
      moveBy: te.moveBy,
    }
  }

  /** Open the paragraph-sized editor over a clustered block: the whole block is the
      edit unit and the commit reflows the text within the block width. Pending plain
      line edits inside the block fold into the draft (the commit then replaces them).
      `fallbackSpan` is the text-layer span under the click: when the dry-run probe
      reports the block cannot be located as one unit (clustering misfires on table
      layouts — vertically stacked cells read as a "paragraph"), committing could only
      ever fail with textEditNoMatch, so degrade to the line-level editor for that
      span instead. */
  const startBlockEdit = (
    origIdx: number,
    block: TextBlock,
    fallbackSpan?: HTMLElement,
    preselect?: [number, number],
  ) => {
    // A pending edit already owns this block (a move, or a block edit whose preview
    // no longer covers the click point): reopen it instead of drafting a duplicate
    const pending = pendingEditFor(origIdx, block)
    if (pending) {
      setSelected(null)
      draftSelectedRef.current = false
      draftPreselectRef.current = null
      setTextDraft(pendingDraft(pending))
      return
    }
    const oldText = joinBlockLines(block.lines.map((l) => l.text))
    if (!oldText.trim()) return
    const rect: [number, number, number, number] = [...block.rect]
    const fold = foldBlockValue(origIdx, block, oldText)
    setSelected(null)
    draftSelectedRef.current = false
    // Preselect offsets are into oldText; folded pending edits shift them
    draftPreselectRef.current = preselect && (fold?.value ?? oldText) === oldText ? preselect : null
    setTextDraft({
      origIdx,
      rect,
      oldText,
      fontSize: block.fontSize,
      value: fold?.value ?? oldText,
      charStyles: fold?.charStyles,
      editId: fold?.editId,
      foldedIds: fold && fold.foldedIds.length > 0 ? fold.foldedIds : undefined,
      foldBase: fold?.value,
      foldStyles: fold?.charStyles,
      block: {
        leftPt: block.rect[0],
        firstBaseline: block.lines[0]!.y,
        widthPt: block.rect[2] - block.rect[0],
        lineHeight: block.lineHeight,
        align: block.align,
        bottomPt: block.rect[1],
      },
    })
    seedDraftFont(origIdx, rect)
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
          if (!v) return
          if (v.reason) {
            // Only swap editors while the draft is untouched — yanking typed text
            // would be worse than the commit-time notice. rect identity pins the
            // draft this probe belongs to (folded drafts carry an editId).
            const d = textDraftRef.current
            if (!d || d.origIdx !== origIdx || d.rect !== rect) return
            if (d.value !== (d.foldBase ?? d.oldText)) return
            setTextDraft(null)
            if (fallbackSpan?.isConnected) openLineEdit(origIdx, fallbackSpan)
            return
          }
          setTextDraft((d) => {
            if (!d || d.origIdx !== origIdx || d.rect !== rect) return d
            let next = d
            if (v.bounds) next = { ...next, cover: v.bounds }
            next = seedDraftColors(next, v)
            return next
          })
        })
        .catch(() => {
          /* cover falls back to the block rect */
        })
    }
  }

  /** A drag over page text in edit mode (WPS-style) opens the editor with exactly the
      dragged characters selected: typing replaces just them, a swatch colors just them.
      Runs off the click that follows the drag's mouseup; returns true when consumed. */
  const dragEditFromSelection = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>): boolean => {
    const sel = window.getSelection()
    if (!sel || sel.isCollapsed || sel.rangeCount === 0) return false
    const range = sel.getRangeAt(0)
    const layer = e.currentTarget.querySelector('.textLayer')
    if (!layer) return false
    const spanOf = (node: Node): HTMLElement | null => {
      const el = node instanceof Element ? node : node.parentElement
      const span = el?.closest('.textLayer span')
      return span instanceof HTMLElement && layer.contains(span) ? span : null
    }
    const offsetIn = (span: HTMLElement, node: Node, off: number) =>
      node === span ? (off === 0 ? 0 : (span.textContent ?? '').length) : off
    const startSpan = spanOf(range.startContainer)
    const endSpan = spanOf(range.endContainer)
    if (!startSpan || !endSpan) return false
    const gs = groupLineSpans(startSpan)
    const ge = gs.spans.includes(endSpan) ? gs : groupLineSpans(endSpan)
    const si = gs.spans.indexOf(startSpan)
    const ei = ge.spans.indexOf(endSpan)
    if (si < 0 || ei < 0 || !gs.text.trim()) return false
    const rawS = gs.starts[si]! + offsetIn(startSpan, range.startContainer, range.startOffset)
    const rawE = ge.starts[ei]! + offsetIn(endSpan, range.endContainer, range.endOffset)
    // The block under the drag's start (same lookup as plain clicks)
    const pageBox = e.currentTarget.getBoundingClientRect()
    const sr = startSpan.getBoundingClientRect()
    const [px, py] = viewToPdf(
      pageGeom(origIdx),
      (sr.left + sr.width / 2 - pageBox.left) / scale,
      (sr.top + sr.height / 2 - pageBox.top) / scale,
    )
    const block = pageBlocks
      .get(origIdx)
      ?.find((b) => px >= b.rect[0] && px <= b.rect[2] && py >= b.rect[1] && py <= b.rect[3])
    if (block && block.lines.length > 1) {
      const blockText = joinBlockLines(block.lines.map((l) => l.text))
      // Cross-line drags map each endpoint through its own visual line
      const pre =
        gs === ge
          ? mapLineRangeToBlock(blockText, gs.text, Math.min(rawS, rawE), Math.max(rawS, rawE))
          : (() => {
              const a = mapLineRangeToBlock(blockText, gs.text, rawS, gs.text.length)
              const b = mapLineRangeToBlock(blockText, ge.text, 0, rawE)
              return a && b
                ? ([Math.min(a[0], b[0]), Math.max(a[1], b[1])] as [number, number])
                : null
            })()
      sel.removeAllRanges()
      startBlockEdit(origIdx, block, startSpan, pre ?? undefined)
      return true
    }
    const pre: [number, number] =
      gs === ge ? [Math.min(rawS, rawE), Math.max(rawS, rawE)] : [rawS, gs.text.length]
    sel.removeAllRanges()
    openLineEdit(origIdx, startSpan, pre)
    return true
  }

  /** Caret of a click on page text: the span under the point plus the click's
      code-unit offset inside its visual line. Unified selection model: a click is a
      zero-length drag, so it carries a collapsed preselect into the opened editor. */
  const caretFromPoint = (e: ReactMouseEvent<HTMLDivElement>) => {
    const layer = e.currentTarget.querySelector('.textLayer')
    const r = document.caretRangeFromPoint(e.clientX, e.clientY)
    if (!layer || !r) return null
    const el =
      r.startContainer instanceof Element ? r.startContainer : r.startContainer.parentElement
    const span = el?.closest('.textLayer span')
    if (!(span instanceof HTMLElement) || !layer.contains(span)) return null
    const group = groupLineSpans(span)
    const i = group.spans.indexOf(span)
    if (i < 0) return null
    const off =
      r.startContainer === span ? 0 : Math.min(r.startOffset, (span.textContent ?? '').length)
    return { span, group, raw: group.starts[i]! + off }
  }

  /** Click on a text-layer span in edit mode → open the floating editor over that run */
  const startTextEdit = (origIdx: number, e: ReactMouseEvent<HTMLDivElement>) => {
    if (!readOnly && dragEditFromSelection(origIdx, e)) {
      e.stopPropagation()
      return
    }
    const span = (e.target as HTMLElement).closest('.textLayer span')
    const caret = caretFromPoint(e)
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
        // Single-line blocks stay on the line path unless a pending block-level
        // edit/move already owns them (startBlockEdit then reopens that edit)
        if (block && (block.lines.length > 1 || pendingEditFor(origIdx, block))) {
          e.stopPropagation()
          const pre = caret
            ? (mapLineRangeToBlock(
                joinBlockLines(block.lines.map((l) => l.text)),
                caret.group.text,
                caret.raw,
                caret.raw,
              ) ?? undefined)
            : undefined
          startBlockEdit(origIdx, block, span instanceof HTMLElement ? span : undefined, pre)
          return
        }
      }
    }
    const anchor = caret?.span ?? (span instanceof HTMLElement ? span : null)
    if (!anchor) return
    if (!(anchor.textContent ?? '').trim()) return
    e.stopPropagation()
    openLineEdit(origIdx, anchor, caret ? [caret.raw, caret.raw] : undefined)
  }

  /** Open the line-level floating editor over the visual line containing `span`.
      Reads live client rects, so it also serves the async block-probe fallback. */
  const openLineEdit = (origIdx: number, span: HTMLElement, preselect?: [number, number]) => {
    const pageEl = span.closest('.pdf-page')
    if (!pageEl) return
    // Edit the whole visual line, not the clicked pdf.js run (CJK is often one span
    // per glyph); the save-side matcher aggregates the covered text objects anyway
    const lineGroup = groupLineSpans(span)
    const oldText = lineGroup.text
    if (!oldText.trim()) return
    const pageBox = pageEl.getBoundingClientRect()
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
    draftPreselectRef.current = preselect ?? null
    setTextDraft({ origIdx, rect, oldText, fontSize, value: oldText })
    seedDraftFont(origIdx, rect)
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
          if (!v) return
          if (v.reason) {
            // The engine cannot locate this line: close the untouched draft with the
            // notice now instead of letting the user edit and fail at commit time.
            // A draft the user already typed into stays open (yanking typed text
            // would be worse) — the commit-time validation still reports it.
            const d = textDraftRef.current
            if (!d || d.editId || d.origIdx !== origIdx || d.rect !== rect) return
            if (d.value !== d.oldText) return
            setTextDraft(null)
            showNotice(t('textEditNoMatch'))
            return
          }
          setTextDraft((d) => {
            if (!d || d.editId || d.origIdx !== origIdx || d.rect !== rect) return d
            let next = d
            if (v.bounds) next = { ...next, cover: v.bounds }
            next = seedDraftColors(next, v)
            return next
          })
        })
        .catch(() => {
          /* cover falls back to the span rect */
        })
    }
  }

  /** A partial-style patch: undefined fields keep, null clears back to inherit */
  type StylePatch = { [K in keyof CharStyle]?: CharStyle[K] | null }

  /** Style-bar action: a partial textarea selection styles just that range (the
      selection survives the control's focus steal); a collapsed caret or a select-all
      applies the whole-draft change and clears the touched fields from any selection
      runs — the whole-draft value owns those fields again (the pre-existing color
      behavior). `refocus` returns the caret to the textarea (skipped for native
      inputs whose panels keep sending changes). */
  const applyDraftStyle = (
    partialPatch: (d: TextDraft, start: number, end: number) => StylePatch,
    whole: (d: TextDraft) => TextDraft,
    refocus: boolean,
  ) => {
    const ta = draftTaRef.current
    setTextDraft((d) => {
      if (!d) return d
      const start = ta?.selectionStart ?? 0
      const end = ta?.selectionEnd ?? 0
      const partial = ta !== null && end > start && !(start === 0 && end >= d.value.length)
      if (!partial) {
        const next = whole(d)
        if (!d.charStyles) return next
        const clear: StylePatch = {}
        for (const k of Object.keys(partialPatch(d, 0, 0)) as (keyof CharStyle)[]) clear[k] = null
        const styles = d.charStyles.map((key) => patchStyle(key, clear))
        return { ...next, charStyles: styles.some((k) => k) ? styles : undefined }
      }
      const patch = partialPatch(d, start, end)
      const styles = d.charStyles ? [...d.charStyles] : Array<string>(d.value.length).fill('')
      for (let i = start; i < end; i++) styles[i] = patchStyle(styles[i]!, patch)
      return { ...d, charStyles: styles.some((k) => k) ? styles : undefined }
    })
    if (refocus && ta) {
      const { selectionStart, selectionEnd } = ta
      ta.focus()
      ta.setSelectionRange(selectionStart, selectionEnd)
    }
  }

  const applyDraftColor = (hex: string, refocus: boolean) =>
    applyDraftStyle(
      () => ({ color: hex }),
      (d) => ({ ...d, color: hex }),
      refocus,
    )

  /** Selection toggle for bold/italic: on unless every selected char is already
      effectively on (explicit override, else the draft-level toggle); a target that
      matches the draft level collapses back to inherit. */
  const toggleDraftFlag = (flag: 'bold' | 'italic') =>
    applyDraftStyle(
      (d, start, end) => {
        const base = d[flag] ?? false
        let all = true
        for (let i = start; i < end && all; i++) {
          all = decodeStyle(d.charStyles?.[i] ?? '')[flag] ?? base
        }
        const target = !all
        return { [flag]: target === base ? null : target }
      },
      (d) => ({ ...d, [flag]: d[flag] ? undefined : true }),
      true,
    )

  /** Fold a floating-editor draft into the pending-edit list; null = nothing changed,
      'overflow' = a block reflow would spill onto occupied space below (not merged —
      the caller notifies and keeps the draft) */
  const mergeTextDraft = (
    edits: LocalTextEdit[],
    d: TextDraft,
  ): LocalTextEdit[] | null | 'overflow' => {
    const existing = d.editId ? edits.find((e) => e.id === d.editId) : undefined
    // Block drafts commit their reflowed form: greedy-wrap each paragraph to the
    // block width in the same face/size the editor previews with, and anchor the
    // rebuilt lines at the block corner with the block's original leading
    let newText = d.value
    let origin: [number, number] | undefined
    let lineLeading: number | undefined
    let lineXOffsets: number[] | undefined
    let overflowed = false
    // An emptied box commits as a deletion: newText '' tells the save to remove the
    // matched run instead of rebuilding it (Escape still discards the draft)
    if (d.value.trim() === '') newText = ''
    if (d.block && d.value.trim() !== '') {
      const size = d.size ?? d.fontSize
      // Measure the reflow in the document's look-alike face when the user
      // didn't pick one — it tracks the save-time advances better than the app font
      const seedF = d.font ? undefined : d.seedFont
      const css =
        (d.font ? EDIT_FONT_BY_ID.get(d.font)?.css : undefined) ??
        seedF?.css ??
        getComputedStyle(document.body).fontFamily
      const cssStyle = `${d.italic || seedF?.italic ? 'italic ' : ''}${
        d.bold ? 'bold' : seedF?.weight ? String(seedF.weight) : ''
      }`.trim()
      lineLeading = d.block.lineHeight * (size / d.fontSize)
      const wrapped = d.value
        .split('\n')
        .flatMap((p) => (p.trim() ? wrapText(p, d.block!.widthPt, size, css, cssStyle) : []))
      // Checked only after the no-op early returns below: a re-wrap in the
      // fallback face can grow a tight block even when nothing was edited,
      // and an untouched dismiss must never be refused
      overflowed = reflowOverflows(
        d.block,
        wrapped.length,
        lineLeading,
        size,
        pageBlocks.get(d.origIdx),
        d.rect,
      )
      newText = wrapped.join('\n')
      origin = [d.block.leftPt, d.block.firstBaseline]
      if (d.block.align !== 'left') {
        lineXOffsets = wrapped.map((l) => {
          const slack = d.block!.widthPt - measurePt(l, size, css, cssStyle)
          return Math.max(0, d.block!.align === 'center' ? slack / 2 : slack)
        })
      }
    }
    // Selection-level styles: draft-space per-char keys carried onto the committed
    // newText (wrapping only rearranges whitespace, so non-ws chars align in order)
    const draftStyles = d.charStyles?.some((c) => c) ? d.charStyles : undefined
    const keyRuns = draftStyles
      ? colorsToRuns(
          newText === d.value ? draftStyles : mapCharColors(d.value, draftStyles, newText),
        )
      : []
    const styleRuns = keyRuns.length ? keyRunsToStyleRuns(keyRuns) : undefined
    const prevValue = existing ? existing.input.newText : d.oldText
    const cmpValue = existing && d.block ? newText : d.value
    const prevSize = existing?.input.newFontSize
    const prevColor = existing?.input.newColor && rgb255ToHex(existing.input.newColor)
    const prevFont = existing?.input.newFont
    const prevBold = existing?.input.newBold ? true : undefined
    const prevItalic = existing?.input.newItalic ? true : undefined
    // Baseline for "did the styles change": the pending edit's committed runs
    // (newText offsets, same space as keyRuns), or — for a fresh draft — the
    // document's own colors the probe seeded (oldText offsets: compare in draft
    // space, the wrap may shift words and move newText offsets without any
    // style changing). Seeded colors alone are not a change — they only ride
    // along so a rebuild repaints them.
    const prevRuns = existing
      ? styleRunsToKeyRuns(existing.input.styleRuns ?? existing.input.colorRuns ?? [])
      : (d.seedStyleRuns ?? [])
    const cmpRuns = existing ? keyRuns : draftStyles ? colorsToRuns(draftStyles) : []
    // Folded paragraph draft committed untouched: keep the folded line edits as
    // they are instead of converting them into a whole-paragraph rebuild (the
    // fold-seeded styles of those edits are not a change either)
    if (
      d.foldBase !== undefined &&
      d.value === d.foldBase &&
      d.size === undefined &&
      d.color === undefined &&
      d.font === undefined &&
      d.bold === undefined &&
      d.italic === undefined &&
      colorRunsEqual(
        draftStyles ? colorsToRuns(draftStyles) : [],
        d.foldStyles ? colorsToRuns(d.foldStyles) : [],
      )
    )
      return null
    if (
      cmpValue === prevValue &&
      d.size === prevSize &&
      d.color === prevColor &&
      d.font === prevFont &&
      d.bold === prevBold &&
      d.italic === prevItalic &&
      colorRunsEqual(cmpRuns, prevRuns)
    )
      return null
    const dropFolded = (list: LocalTextEdit[]) =>
      d.foldedIds && d.foldedIds.length > 0
        ? list.filter((e) => !d.foldedIds!.includes(e.id))
        : list
    // A pure move reopened and closed untouched must stay a pure move: its stored
    // newText keeps the document's own line breaks (which a DOM re-wrap would not
    // reproduce), and "unchanged text" must not read as "revert" below — that
    // would silently delete the pending move
    if (
      existing?.input.translate &&
      d.value === d.oldText &&
      d.size === undefined &&
      d.color === undefined &&
      d.font === undefined &&
      d.bold === undefined &&
      d.italic === undefined &&
      keyRuns.length === 0
    )
      return null
    if (
      d.value === d.oldText &&
      d.size === undefined &&
      d.color === undefined &&
      d.font === undefined &&
      d.bold === undefined &&
      d.italic === undefined &&
      keyRuns.length === 0
    ) {
      // Reverted back to the original — the pending edit(s) are moot
      return dropFolded(edits.filter((e) => e.id !== d.editId))
    }
    if (overflowed) return 'overflow'
    const input: TextEditInput = {
      pageIndex: d.origIdx,
      rect: d.rect,
      oldText: d.oldText,
      newText,
      fontSize: d.fontSize,
      newFontSize: d.size,
      newColor: d.color === undefined ? undefined : hexTo255(d.color),
      styleRuns,
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
      ? dropFolded(edits).map((e) =>
          e.id === d.editId
            ? {
                ...e,
                input,
                cover: d.cover ?? e.cover,
                baseInk: d.seedInk ?? e.baseInk,
                baseFont: d.seedFont ?? e.baseFont,
              }
            : e,
        )
      : [...edits, { id: newId(), input, cover: d.cover, baseInk: d.seedInk, baseFont: d.seedFont }]
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
          showNotice(t(edit.input.translate ? 'textBlockMoveNoMatch' : 'textEditNoMatch'))
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
    const merged = mergeTextDraft(textEdits, d)
    if (merged === 'overflow') {
      // Keep the editor open so the user can shorten the text (or Escape out)
      showNotice(t('textBlockOverflow'))
      return textEdits
    }
    setTextDraft(null)
    if (!merged) return textEdits
    pushUndo()
    setTextEdits(merged)
    // New edits append; re-opened ones keep their id
    const committed = d.editId ? merged.find((e) => e.id === d.editId) : merged[merged.length - 1]
    if (committed) validateTextEdit(committed)
    return merged
  }

  /** Editor trash button / backspace-on-empty: delete the whole run the editor
      holds, as if the user cleared the text and committed */
  const deleteDraftRun = () => {
    const d = textDraft
    if (!d) return
    setTextDraft(null)
    const merged = mergeTextDraft(textEdits, { ...d, value: '' })
    // A deletion ('' skips the block reflow) can never overflow
    if (!merged || merged === 'overflow') return
    pushUndo()
    setTextEdits(merged)
    const committed = d.editId ? merged.find((e) => e.id === d.editId) : merged[merged.length - 1]
    if (committed) validateTextEdit(committed)
  }

  /** Commit a border-drag of a clustered block as a pending move (WPS-style).
      An untouched block becomes a pure-translate edit — the engine shifts the
      original text objects as-is, so fonts/kerning/leading survive byte-identical.
      A block already carrying a pending edit shifts that edit's position instead,
      and pending plain line edits inside the block fold into one moved paragraph
      rebuild (two edits addressing the same objects would collide at save). */
  const commitBlockMove = (origIdx: number, block: TextBlock, d: [number, number]) => {
    /** Shift one pending edit by d: block rebuilds move their origin, pure moves
        their translate; a line edit carries no position of its own (the rebuild
        sits at the matched objects), so moving it converts it to an origin-anchored
        rebuild at its own shifted line start — restyled/edited lines cannot take
        the pure-translate path, that would discard their pending changes. */
    const shiftEdit = (existing: LocalTextEdit) => {
      pushUndo()
      setTextEdits((prev) =>
        prev.map((te) => {
          if (te.id !== existing.id) return te
          const input = { ...te.input }
          if (input.origin) {
            input.origin = [input.origin[0] + d[0], input.origin[1] + d[1]]
          } else if (!input.translate) {
            // Anchor at the edit's own rect, not the block corner: the edit's
            // visual line can start left of the dragged block (the DOM line
            // grouping joins runs the clustering split off). Baseline comes from
            // the block row containing the edit; buildLine puts the row bottom
            // 0.2 font sizes under the baseline, hence the fallback.
            const cy = (input.rect[1] + input.rect[3]) / 2
            const row = block.lines.find((l) => cy >= l.rect[1] && cy <= l.rect[3])
            input.origin = [
              input.rect[0] + d[0],
              (row?.y ?? input.rect[1] + input.fontSize * 0.2) + d[1],
            ]
            input.lineLeading ??= block.lineHeight
          }
          if (input.translate)
            input.translate = [input.translate[0] + d[0], input.translate[1] + d[1]]
          return {
            ...te,
            input,
            moveBy: [(te.moveBy?.[0] ?? 0) + d[0], (te.moveBy?.[1] ?? 0) + d[1]],
          }
        }),
      )
    }
    const existing = pendingEditFor(origIdx, block)
    if (existing) {
      shiftEdit(existing)
      return
    }
    const oldText = joinBlockLines(block.lines.map((l) => l.text))
    if (!oldText.trim()) return
    const fold = foldBlockValue(origIdx, block, oldText)
    if (fold) {
      const draft: TextDraft = {
        origIdx,
        rect: [...block.rect],
        oldText,
        fontSize: block.fontSize,
        value: fold.value,
        charStyles: fold.charStyles,
        editId: fold.editId,
        foldedIds: fold.foldedIds.length > 0 ? fold.foldedIds : undefined,
        block: {
          leftPt: block.rect[0] + d[0],
          firstBaseline: block.lines[0]!.y + d[1],
          widthPt: block.rect[2] - block.rect[0],
          lineHeight: block.lineHeight,
          align: block.align,
          bottomPt: block.rect[1] + d[1],
        },
      }
      const merged = mergeTextDraft(textEdits, draft)
      if (merged === 'overflow') {
        showNotice(t('textBlockOverflow'))
        return
      }
      if (!merged) return
      pushUndo()
      const withMove = merged.map((te): LocalTextEdit => {
        return te.id === fold.editId ? { ...te, moveBy: d } : te
      })
      setTextEdits(withMove)
      const committed = withMove.find((te) => te.id === fold.editId)
      if (committed) validateTextEdit(committed)
      return
    }
    // The text matchers can miss even though a pending edit claims this block's
    // objects: a DOM visual line can join runs the clustering buckets into a
    // neighboring block, so the edit's oldText covers more than the block and no
    // text comparison lines up. A fresh translate edit would collide with it at
    // save (one of the two silently skipped) — geometric ownership decides
    // instead, and the whole edited line moves with the block.
    const overlapping = blockOverlapEdit(origIdx, block)
    if (overlapping) {
      shiftEdit(overlapping)
      return
    }
    pushUndo()
    const input: TextEditInput = {
      pageIndex: origIdx,
      rect: [...block.rect],
      oldText,
      // The document's own visual lines, so the pending preview stacks them the
      // way the page draws them (the engine never rebuilds a pure move)
      newText: block.lines.map((l) => l.text).join('\n'),
      fontSize: block.fontSize,
      origin: [block.rect[0] + d[0], block.lines[0]!.y + d[1]],
      lineLeading: block.lineHeight,
      align: block.align !== 'left' ? block.align : undefined,
      blockSource: oldText,
      translate: d,
    }
    const te: LocalTextEdit = { id: newId(), input, moveBy: d }
    setTextEdits((prev) => [...prev, te])
    validateTextEdit(te)
  }

  const deleteSelected = () => {
    const sel = selected
    if (!sel) return
    pushUndo()
    if (sel.kind === 'markup') setMarkups((prev) => prev.filter((m) => m.id !== sel.id))
    else if (sel.kind === 'savedMarkup')
      setAnnotDeletes((prev) => [...prev, { id: newId(), annot: sel.annot }])
    else if (sel.kind === 'drawing') setDrawings((prev) => prev.filter((d) => d.id !== sel.id))
    else if (sel.kind === 'textEdit') setTextEdits((prev) => prev.filter((e) => e.id !== sel.id))
    else if (sel.kind === 'textInsert')
      setTextInserts((prev) => prev.filter((insert) => insert.id !== sel.id))
    else if (sel.kind === 'imageEdit')
      setImageEdits((prev) =>
        prev.flatMap((edit) => {
          if (edit.id !== sel.id) return [edit]
          if (
            edit.staticFill &&
            (edit.input.kind === 'transformImage' || edit.input.kind === 'replaceImage')
          ) {
            return [
              {
                ...edit,
                input: {
                  kind: 'deleteImage' as const,
                  pageIndex: edit.input.pageIndex,
                  oldRect: edit.input.oldRect,
                },
              },
            ]
          }
          return []
        }),
      )
    else if (sel.kind === 'pageImage')
      // Deleting an untouched existing image = a pending delete op
      setImageEdits((prev) => [
        ...prev,
        {
          id: newId(),
          input: { kind: 'deleteImage', pageIndex: sel.ref.pageIndex, oldRect: sel.ref.rect },
          staticFill: savedStaticFillForRef(sel.ref),
        },
      ])
    else setStampCfg(null)
    setSelected(null)
    // Transient "deleted · undo" toast so the removal is visible and reversible in place
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current)
    setDeletedInsertedText(sel.kind === 'textInsert')
    setDeleteToast(true)
    toastTimerRef.current = window.setTimeout(() => setDeleteToast(false), 5000)
  }

  /** Show a transient toast; the save-failure badge alone hides the actual reason */
  const showNotice = (msg: string) => {
    if (noticeTimerRef.current !== null) window.clearTimeout(noticeTimerRef.current)
    setNotice(msg)
    noticeTimerRef.current = window.setTimeout(() => setNotice(null), 8000)
  }

  useEffect(() => {
    if (!formHasXfa || !filePath || warnedXfaPathRef.current === filePath) return
    warnedXfaPathRef.current = filePath
    showNotice(t('formXfaWarning'))
  }, [filePath, formHasXfa])

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

  const noticeSkippedTextInserts = (skipped: TextInsertFailure[]) => {
    const pages = [...new Set(skipped.map((s) => s.pageIndex + 1))].sort((a, b) => a - b).join(', ')
    showNotice(t('textInsertSkipped', { pages }))
  }

  /** Pending edits in SavePdfRequest form; shared by in-place Save and Save As.
      `noteFlush` carries the drawings/noteEdits returned by commitNoteEdit — the
      state values in this closure predate that flush. */
  const editsPayload = (
    edits: LocalTextEdit[] = textEdits,
    noteFlush?: { drawings: LocalDrawing[]; noteEdits: LocalNoteEdit[] },
  ) => ({
    markups: markups.map(({ id: _id, ...rest }) => rest),
    annotDeletes: annotDeletes.map((d): AnnotDeleteInput => ({
      pageIndex: d.annot.pageIndex,
      objNum: d.annot.objNum,
      subtype: d.annot.type,
      rect: d.annot.rect,
      // A note thread's comments all share the root's rect; contents disambiguates
      ...(d.annot.type === 'note' ? { contents: d.annot.contents } : {}),
    })),
    noteEdits: (noteFlush?.noteEdits ?? noteEdits).map((e): NoteEditInput => ({
      pageIndex: e.annot.pageIndex,
      objNum: e.annot.objNum,
      rect: e.annot.rect,
      oldContents: e.annot.contents,
      contents: e.contents,
    })),
    drawings: (noteFlush?.drawings ?? drawings).map((d) => d.input),
    textEdits: edits.map((e) => e.input),
    textInserts: textInserts.map((insert) => insert.input),
    imageEdits: imageEdits.map((e) => e.input),
    staticFormFills,
    stamps: stampCfg ? renderStamps(stampCfg, visList) : [],
    formValues: [...formEdits.values()],
    rotations: [...rotations].map(([pageIndex, delta]) => ({ pageIndex, delta })),
    deletedPages: [...deleted],
    ...(order ? { pageOrder: visList } : {}),
    ...(metadata ? { metadata } : {}),
  })

  /** Resolved when the running save() lands; queued saves and Save As serialize behind it */
  const saveInFlightRef = useRef<Promise<boolean> | null>(null)
  /** objNum → /Contents the in-flight save is writing via noteEdits. appliedNoteEdit
      consults it so a re-edit during the write compares against what the file is about
      to say, not the pre-save text (dropping on the wrong compare loses the edit). */
  const inFlightNoteWritesRef = useRef<Map<number, string>>(new Map())
  /** The in-flight save's page remap (old original index → index in the saved file);
      Save As addresses late note edits through it once that save has rewritten pages */
  const inFlightPageMapRef = useRef<Map<number, number> | null>(null)
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
    // Same for an open comment-edit box in the notes margin
    const noteFlush = commitNoteEdit()
    const anythingToSave =
      dirty ||
      edits !== textEdits ||
      noteFlush.drawings !== drawings ||
      noteFlush.noteEdits !== noteEdits
    if (!anythingToSave || !filePath) return Promise.resolve(!anythingToSave)
    // An explicit save opts this file into autosave
    if (!autosave) savedOnceRef.current = true
    // What this save writes — the post-save reload subtracts exactly this, keeping
    // any edits the user makes while the write is in flight
    const snapshot: SavedSnapshot = {
      markupIds: new Set(markups.map((mk) => mk.id)),
      annotDeleteIds: new Set(annotDeletes.map((d) => d.id)),
      noteEditIds: new Set(noteFlush.noteEdits.map((e) => e.id)),
      noteEditWritten: new Map(noteFlush.noteEdits.map((e) => [e.annot.objNum, e.contents])),
      drawingIds: new Set(noteFlush.drawings.map((dr) => dr.id)),
      textEditIds: new Set(edits.map((te) => te.id)),
      textInsertIds: new Set(textInserts.map((insert) => insert.id)),
      imageEditIds: new Set(imageEdits.map((ie) => ie.id)),
      stampCfg,
      formEdits,
      rotations,
      metadata,
      pageMap: new Map(visList.map((origIdx, i) => [origIdx, i])),
    }
    inFlightNoteWritesRef.current = snapshot.noteEditWritten
    inFlightPageMapRef.current = snapshot.pageMap
    const run = (async (): Promise<boolean> => {
      setSaveState('saving')
      const result = await window.pdfApi.save({ path: filePath, ...editsPayload(edits, noteFlush) })
      if (!result.ok) {
        opFailed(result.error)
        return false
      }
      if (result.skippedTextEdits && result.skippedTextEdits.length > 0) {
        noticeSkippedEdits(result.skippedTextEdits)
      }
      if (result.skippedTextInserts && result.skippedTextInserts.length > 0) {
        noticeSkippedTextInserts(result.skippedTextInserts)
      }
      if (result.skippedImageEdits && result.skippedImageEdits.length > 0) {
        noticeSkippedImages(result.skippedImageEdits)
      }
      // Reload: changes are in the file now, canvas renders directly, saved pending ops are cleared
      try {
        const el = scrollRef.current
        const scrollTop = el?.scrollTop ?? 0
        // Page structure/rotation changes cannot retain their old overlays because
        // their geometry or page numbers no longer match the newly saved document.
        const canRetainPreview = rotations.size === 0 && deleted.size === 0 && order === null
        const renderedPageNos = canRetainPreview
          ? [...visibleRows].flatMap((rowIdx) => (rows[rowIdx] ?? []).map((origIdx) => origIdx + 1))
          : []
        await loadDoc(filePath, doc, snapshot, renderedPageNos)
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
      if (saveInFlightRef.current === tracked) {
        saveInFlightRef.current = null
        inFlightNoteWritesRef.current = new Map()
        inFlightPageMapRef.current = null
      }
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
    // The copy must contain what the user sees, including an open floating-editor
    // draft and an open comment-edit box
    const draftEdits = commitTextDraft()
    const noteFlush = commitNoteEdit()
    // A save already in flight (autosave that started before the dialog opened) lands
    // first. If it succeeded, every edit that was pending is now part of the source
    // bytes, so the copy applies nothing on top — deriving this from the save result
    // (instead of re-reading state) avoids racing React's render of the cleared edits.
    const inFlight = saveInFlightRef.current
    // What that save is writing per note, and its page remap — captured now, the
    // refs clear when it lands
    const inFlightWrites = new Map(inFlightNoteWritesRef.current)
    const inFlightPageMap = inFlightPageMapRef.current ? new Map(inFlightPageMapRef.current) : null
    const flushed = inFlight ? await inFlight.catch(() => false) : false
    // The comment edit committed by commitNoteEdit above started after the in-flight
    // save built its payload, so "everything was flushed" never covers it: the copy
    // still needs that rewrite, matched against what the flushed save left on disk.
    const preFlushIds = new Set(noteEdits.map((e) => e.id))
    const lateNoteEdits = flushed
      ? noteFlush.noteEdits
          .filter((e) => !preFlushIds.has(e.id))
          .flatMap((e): NoteEditInput[] => {
            // The flushed save may have deleted or reordered pages: address the note
            // by its index in the rewritten file; its page being gone drops the edit
            const ni = inFlightPageMap ? inFlightPageMap.get(e.annot.pageIndex) : e.annot.pageIndex
            if (ni === undefined) return []
            return [
              {
                pageIndex: ni,
                objNum: e.annot.objNum,
                rect: e.annot.rect,
                oldContents: inFlightWrites.get(e.annot.objNum) ?? e.annot.contents,
                contents: e.contents,
              },
            ]
          })
      : []
    const edits = flushed
      ? {
          markups: [],
          drawings: [],
          formValues: [],
          stamps: [],
          textEdits: [],
          textInserts: [],
          imageEdits: [],
          ...(lateNoteEdits.length > 0 ? { noteEdits: lateNoteEdits } : {}),
        }
      : editsPayload(draftEdits, noteFlush)
    setSaveState('saving')
    const result = await window.pdfApi.save({ path: filePath, targetPath, ...edits })
    if (!result.ok) {
      opFailed(result.error)
      return false
    }
    if (result.skippedTextEdits && result.skippedTextEdits.length > 0) {
      noticeSkippedEdits(result.skippedTextEdits)
    }
    if (result.skippedTextInserts && result.skippedTextInserts.length > 0) {
      noticeSkippedTextInserts(result.skippedTextInserts)
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

  const rotatePages = (origIdxs: number[], dir: 90 | -90) => {
    if (readOnly || origIdxs.length === 0) return
    pushUndo()
    const pages = new Set(origIdxs)
    setRotations((prev) => {
      const next = new Map(prev)
      for (const origIdx of pages) {
        const nv = ((next.get(origIdx) ?? 0) + dir + 360) % 360
        if (nv === 0) next.delete(origIdx)
        else next.set(origIdx, nv)
      }
      return next
    })
    // Image stamps are always drawn upright (both in the overlay and in the saved
    // appearance), so a 90° page turn swaps their displayed width/height. Swap the
    // user-space rect around its center to keep the bitmap's aspect ratio intact.
    setDrawings((prev) =>
      prev.map((d) => {
        if (d.input.kind !== 'image' || !pages.has(d.input.pageIndex)) return d
        const [x1, y1, x2, y2] = d.input.rect
        const cx = (x1 + x2) / 2
        const cy = (y1 + y2) / 2
        const hw = (x2 - x1) / 2
        const hh = (y2 - y1) / 2
        return { ...d, input: { ...d.input, rect: [cx - hh, cy - hw, cx + hh, cy + hw] } }
      }),
    )
  }

  const rotatePage = (origIdx: number, dir: 90 | -90) => rotatePages([origIdx], dir)

  const rotateAllPages = (dir: 90 | -90) => rotatePages(visList, dir)

  /** Reverse the visible page order; deleted pages stay at the tail like movePage */
  const reversePages = () => {
    if (pageCount <= 1 || readOnly) return
    pushUndo()
    const next = [...visList].reverse()
    const rest = sizes.map((_, i) => i).filter((i) => !next.includes(i))
    setOrder([...next, ...rest])
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

  /** Fit a visual signature into an AcroForm /Sig widget. */
  const placeSignatureInField = (sig: SignatureData, target: FormWidget) => {
    pushUndo()
    setDrawings((prev) => [
      ...prev,
      {
        id: newId(),
        input: signatureDrawingForField(sig, target, drawColor),
        formWidgetId: target.id,
      },
    ])
    setSignatureTarget(null)
    setActiveFormWidgetId(target.id)
  }

  // ── Image editing (content-stream image ops, applied by the main process at save) ──

  const prepareStaticFormFill = (
    kind: StaticFormFillKind,
    image: Extract<SignatureData, { kind: 'image' }>,
  ) => {
    setEditTextMode(false)
    setTextDraft(null)
    setPendingTextInsert(null)
    setDrawTool(null)
    setPendingSign(null)
    setSignatureTarget(null)
    setEditImageMode(false)
    setPendingStaticFill(kind)
    setImagePick(image)
  }

  const startStaticFormMark = (kind: 'check' | 'cross') => {
    if (pendingStaticFill === kind) {
      setImagePick(null)
      setPendingStaticFill(null)
      return
    }
    const image = renderStaticFormMark(kind)
    if (image) prepareStaticFormFill(kind, image)
  }

  const textInsertOffsets = (
    text: string,
    fontSize: number,
    align: 'left' | 'center' | 'right',
  ): number[] => {
    if (align === 'left') return text.split('\n').map(() => 0)
    const font = `${fontSize}px ${getComputedStyle(document.body).fontFamily}`
    return text
      .split('\n')
      .map((line) => measureTextWidth(line, font) * (align === 'center' ? -0.5 : -1))
  }

  const confirmTextInsert = () => {
    const text = staticText.trim()
    if (!text) return
    const config: Omit<TextInsertInput, 'pageIndex' | 'origin'> = {
      text,
      fontSize: staticTextSize,
      color: hexTo255(staticTextColor),
      lineLeading: staticTextSize * 1.2,
      lineXOffsets: textInsertOffsets(text, staticTextSize, staticTextAlign),
      align: staticTextAlign,
    }
    setStaticTextDialog(false)
    if (textInsertEditId) {
      pushUndo()
      setTextInserts((prev) =>
        prev.map((insert) =>
          insert.id === textInsertEditId
            ? { ...insert, input: { ...insert.input, ...config } }
            : insert,
        ),
      )
      setTextInsertEditId(null)
      return
    }
    setPendingTextInsert(config)
  }

  const confirmStaticFormText = () => {
    if (staticTextPurpose === 'insert') {
      confirmTextInsert()
      return
    }
    const image = renderStaticFormText(staticText, staticTextSize, staticTextColor, staticTextAlign)
    if (!image) return
    setStaticTextDialog(false)
    const target = staticTextEditTarget
    if (target) {
      const updated: StaticFormFillRecord = {
        ...target.record,
        text: staticText,
        fontSize: staticTextSize,
        color: staticTextColor,
        align: staticTextAlign,
      }
      setStaticTextEditTarget(null)
      if (target.kind === 'saved') {
        replaceExisting(target.ref, image.image, updated)
      } else {
        pushUndo()
        setSelected(null)
        setImageEdits((prev) =>
          prev.map((edit) => {
            if (edit.id !== target.editId || edit.input.kind === 'deleteImage') return edit
            const input =
              edit.input.kind === 'insertImage'
                ? { ...edit.input, image: image.image }
                : edit.input.kind === 'replaceImage'
                  ? { ...edit.input, image: image.image }
                  : {
                      kind: 'replaceImage' as const,
                      pageIndex: edit.input.pageIndex,
                      oldRect: edit.input.oldRect,
                      rect: edit.input.rect,
                      image: image.image,
                      layer: edit.input.layer,
                      quarterTurns: edit.input.quarterTurns,
                    }
            return { ...edit, input, staticFill: { ...updated, rect: input.rect } }
          }),
        )
      }
      return
    }
    prepareStaticFormFill('text', image)
  }

  const placeTextInsert = (origIdx: number, vx: number, vy: number) => {
    const pending = pendingTextInsert
    if (!pending) return
    const geom = pageGeom(origIdx)
    pushUndo()
    setTextInserts((prev) => [
      ...prev,
      {
        id: newId(),
        input: {
          ...pending,
          pageIndex: origIdx,
          origin: viewToPdf(geom, vx, vy),
          rotate: ((geom.rot % 360) + 360) % 360,
        },
      },
    ])
    setPendingTextInsert(null)
    setTextInsertPointer(null)
  }

  /** Ribbon button → file picker; the picked image then enters click-to-place mode */
  const pickInsertImage = () => {
    setEditTextMode(false)
    setTextDraft(null)
    setPendingTextInsert(null)
    setDrawTool(null)
    setPendingSign(null)
    setSignatureTarget(null)
    setPendingStaticFill(null)
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
      commitBaked(target, base64)
      return
    }
    setImagePick({ kind: 'image', image: base64, width: canvas.width, height: canvas.height })
    setPendingStaticFill(null)
  }

  /** Drop the picked image centered on the click point, into the text-below band by default */
  const placeImage = (origIdx: number, vx: number, vy: number) => {
    const pick = imagePick
    if (!pick) return
    const geom = pageGeom(origIdx)
    const disp = geomDispSize(geom)
    const k = pendingStaticFill
      ? staticFormFillPlaceK()
      : imagePlaceK(pick, disp.width, disp.height)
    const w = pick.width * k
    const h = pick.height * k
    const left = Math.min(Math.max(vx - w / 2, 0), Math.max(disp.width - w, 0))
    const top = Math.min(Math.max(vy - h / 2, 0), Math.max(disp.height - h, 0))
    const [ax, ay] = viewToPdf(geom, left, top)
    const [bx, by] = viewToPdf(geom, left + w, top + h)
    const rect: [number, number, number, number] = [
      Math.min(ax, bx),
      Math.min(ay, by),
      Math.max(ax, bx),
      Math.max(ay, by),
    ]
    const id = newId()
    const staticFill: StaticFormFillRecord | undefined = pendingStaticFill
      ? {
          id,
          kind: pendingStaticFill,
          pageIndex: origIdx,
          rect,
          ...(pendingStaticFill === 'text'
            ? {
                text: staticText,
                fontSize: staticTextSize,
                color: staticTextColor,
                align: staticTextAlign,
              }
            : {}),
        }
      : undefined
    pushUndo()
    setImageEdits((prev) => [
      ...prev,
      {
        id,
        input: {
          kind: 'insertImage',
          pageIndex: origIdx,
          image: pick.image,
          rect,
          layer: pendingStaticFill ? 'aboveText' : 'belowText',
          rotate: ((geom.rot % 360) + 360) % 360,
        },
        staticFill,
      },
    ])
    setImagePick(null)
    setPendingStaticFill(null)
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

  const savedStaticFillForRef = (ref: PageImageRef): StaticFormFillRecord | undefined =>
    savedStaticFormFills.find(
      (record) => record.pageIndex === ref.pageIndex && rectsNear(record.rect, ref.rect),
    )

  const selectedStaticTextTarget = () => {
    if (selected?.kind === 'pageImage') {
      const record = savedStaticFillForRef(selected.ref)
      return record?.kind === 'text'
        ? ({ kind: 'saved', ref: selected.ref, record } as const)
        : null
    }
    if (selected?.kind === 'imageEdit') {
      const edit = imageEdits.find((candidate) => candidate.id === selected.id)
      return edit?.staticFill?.kind === 'text'
        ? ({ kind: 'pending', editId: edit.id, record: edit.staticFill } as const)
        : null
    }
    return null
  }

  const startEditStaticText = () => {
    const target = selectedStaticTextTarget()
    if (!target) return
    setStaticText(target.record.text ?? '')
    setStaticTextSize(target.record.fontSize ?? 14)
    setStaticTextColor(target.record.color ?? '#111111')
    setStaticTextAlign(target.record.align ?? 'left')
    setStaticTextEditTarget(target)
    setStaticTextPurpose('form')
    setTextInsertEditId(null)
    setSelected(null)
    setStaticTextDialog(true)
  }

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
    const savedStaticFill = savedStaticFillForRef(ref)
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
        staticFill: savedStaticFill ? { ...savedStaticFill, rect } : undefined,
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

  /** Rotate PNG pixels by 0-3 screen-clockwise quarter turns (fresh inserts carry
      rotation baked into the bytes; the bake pipeline collapses pending op turns) */
  const rotatePngTurns = (b64: string, turns: number): Promise<string | null> => {
    const tn = ((turns % 4) + 4) % 4
    if (tn === 0) return Promise.resolve(b64)
    return new Promise((resolve) => {
      const img = new Image()
      img.onload = () => {
        const c = document.createElement('canvas')
        c.width = tn % 2 === 0 ? img.width : img.height
        c.height = tn % 2 === 0 ? img.height : img.width
        const ctx = c.getContext('2d')
        if (!ctx) return resolve(null)
        ctx.translate(c.width / 2, c.height / 2)
        ctx.rotate((tn * 90 * Math.PI) / 180)
        ctx.drawImage(img, -img.width / 2, -img.height / 2)
        resolve(c.toDataURL('image/png').split(',')[1] ?? null)
      }
      img.onerror = () => resolve(null)
      img.src = `data:image/png;base64,${b64}`
    })
  }

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
      void rotatePngTurns(image, turn).then((rotated) => {
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
              ? {
                  ...e,
                  // rotating the bytes invalidates a recorded pre-transparency base
                  opacityBase: undefined,
                  input: { ...e.input, image: rotated, rect: rotatedRect(e.input.rect) },
                }
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
  const replaceExisting = (ref: PageImageRef, png: string, staticFill?: StaticFormFillRecord) => {
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
        staticFill,
      },
    ])
  }

  // ── Baked pixel edits (flip / transparency / crop / cutout / replace) ──
  // PDF images are bitmaps, so these all land the same way: fetch the selection's
  // displayed pixels, transform them on a canvas, and write the result back — a new
  // replaceImage op for an untouched image, or an in-place byte swap of a pending op
  // (transform ops morph into replaceImage with their quarter turns baked in).

  /** What a pixel edit applies to; `before` guards edit ops against concurrent changes
      (state inputs are immutable, so reference equality detects any interim mutation) */
  type ImageBakeTarget =
    { kind: 'existing'; ref: PageImageRef } | { kind: 'edit'; id: string; before: ImageEditInput }

  const bakeTargetOf = (sel: AnnotSelection | null): ImageBakeTarget | null => {
    if (sel?.kind === 'pageImage') return { kind: 'existing', ref: sel.ref }
    if (sel?.kind !== 'imageEdit') return null
    const e = imageEdits.find((x) => x.id === sel.id)
    if (!e || e.input.kind === 'deleteImage') return null
    return { kind: 'edit', id: sel.id, before: e.input }
  }

  /** Upscale over the on-page size when fetching bake sources, so crop/flip results
      stay print-sharp instead of inheriting the ~1px/pt ghost resolution */
  const BAKE_SCALE = 3

  /** The target's pixels in displayed orientation (pending quarter turns baked in) */
  const bakeSourcePng = async (target: ImageBakeTarget): Promise<string | null> => {
    const fetchPng = (pageIndex: number, rect: [number, number, number, number]) =>
      window.pdfApi
        .pageImagePng({ path: filePath, pageIndex, rect, scale: BAKE_SCALE })
        .catch(() => null)
    if (target.kind === 'existing') return fetchPng(target.ref.pageIndex, target.ref.rect)
    const input = target.before
    if (input.kind === 'insertImage') return input.image
    if (input.kind === 'replaceImage') return rotatePngTurns(input.image, input.quarterTurns ?? 0)
    if (input.kind === 'transformImage') {
      const src = await fetchPng(input.pageIndex, input.oldRect)
      return src ? rotatePngTurns(src, input.quarterTurns ?? 0) : null
    }
    return null
  }

  /** Decode a PNG, run a same-size pixel transform, re-encode (null on failure) */
  const transformPngPixels = (
    b64: string,
    fn: (img: PixelImage) => Uint8ClampedArray<ArrayBuffer>,
  ): Promise<string | null> =>
    new Promise((resolve) => {
      const img = new Image()
      img.onload = () => {
        const w = img.naturalWidth
        const h = img.naturalHeight
        const c = document.createElement('canvas')
        c.width = w
        c.height = h
        const ctx = c.getContext('2d')
        if (!ctx || !w || !h) return resolve(null)
        ctx.drawImage(img, 0, 0)
        const d = ctx.getImageData(0, 0, w, h)
        ctx.putImageData(new ImageData(fn({ data: d.data, width: w, height: h }), w, h), 0, 0)
        resolve(c.toDataURL('image/png').split(',')[1] ?? null)
      }
      img.onerror = () => resolve(null)
      img.src = `data:image/png;base64,${b64}`
    })

  /** Crop footprint of an insert op: its bytes are display-oriented, so the fractions
      apply in display space and map back through the page geometry (handles /Rotate) */
  const cropRectDisplay = (
    pageIndex: number,
    rect: [number, number, number, number],
    crop: CropFractions,
  ): [number, number, number, number] => {
    const geom = pageGeom(pageIndex)
    const box = pdfRectToCss(geom, rect, 1)
    const [ax, ay] = viewToPdf(geom, box.left + crop.l * box.width, box.top + crop.t * box.height)
    const [bx, by] = viewToPdf(geom, box.left + crop.r * box.width, box.top + crop.b * box.height)
    return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)]
  }

  /** Land baked pixels on the target (see the section comment); crop also shrinks the
      footprint to the kept region. Silently drops results that lost a race.
      opacityBase marks the bytes as a transparency bake (see LocalImageEdit); any other
      pixel edit leaves it unset, which clears a stale base. */
  const commitBaked = (
    target: ImageBakeTarget,
    png: string,
    crop?: CropFractions,
    opacityBase?: string,
  ) => {
    if (target.kind === 'existing') {
      const ref = target.ref
      const key = `${ref.pageIndex}:${imageRectKey(ref.rect)}`
      const claimed = imageEditsRef.current.some(
        (e) =>
          e.input.kind !== 'insertImage' &&
          `${e.input.pageIndex}:${imageRectKey(e.input.oldRect)}` === key,
      )
      if (claimed) return
      pushUndoRef.current()
      setImageEdits((prev) => [
        ...prev,
        {
          id: newId(),
          input: {
            kind: 'replaceImage',
            pageIndex: ref.pageIndex,
            oldRect: ref.rect,
            rect: crop ? cropRect(ref.rect, crop) : ref.rect,
            image: png,
          },
          origAbove: ref.aboveText,
          opacityBase,
        },
      ])
      return
    }
    const cur = imageEditsRef.current.find((x) => x.id === target.id)
    if (!cur || cur.input !== target.before || cur.input.kind === 'deleteImage') return
    pushUndoRef.current()
    setImageEdits((prev) =>
      prev.map((e) => {
        if (e.id !== target.id || e.input.kind !== cur.input.kind) return e
        if (e.input.kind === 'insertImage') {
          return {
            ...e,
            opacityBase,
            input: {
              ...e.input,
              image: png,
              rect: crop ? cropRectDisplay(e.input.pageIndex, e.input.rect, crop) : e.input.rect,
            },
          }
        }
        if (e.input.kind !== 'transformImage' && e.input.kind !== 'replaceImage') return e
        return {
          ...e,
          opacityBase,
          input: {
            kind: 'replaceImage',
            pageIndex: e.input.pageIndex,
            oldRect: e.input.oldRect,
            rect: crop ? cropRect(e.input.rect, crop) : e.input.rect,
            image: png,
            ...(e.input.layer ? { layer: e.input.layer } : {}),
          },
        }
      }),
    )
  }

  /** Fetch → transform → commit, used by the one-click bakes (flip / transparency) */
  const bakeSelected = (fn: (img: PixelImage) => Uint8ClampedArray<ArrayBuffer>) => {
    const target = bakeTargetOf(selected)
    if (!target) return
    setSelected(null)
    void (async () => {
      const src = await bakeSourcePng(target)
      const out = src ? await transformPngPixels(src, fn) : null
      if (out) commitBaked(target, out)
    })()
  }

  const flipSelected = (axis: 'h' | 'v') => bakeSelected((img) => flipPixels(img, axis))

  /** Bake a transparency preset (percent transparent, slides-style ladder). Absolute,
      not compounding: when the op's bytes came from an earlier transparency bake,
      re-bake from the recorded pre-transparency pixels (bytes-only swap — rect, layer
      and pending quarter turns stay put). */
  const applyImageOpacity = (pct: number) => {
    const target = bakeTargetOf(selected)
    if (!target) return
    setSelected(null)
    const alpha = (img: PixelImage) => multiplyAlpha(img, 1 - pct / 100)
    void (async () => {
      const prior =
        target.kind === 'edit'
          ? imageEditsRef.current.find((x) => x.id === target.id)?.opacityBase
          : undefined
      if (prior && target.kind === 'edit') {
        const out = await transformPngPixels(prior, alpha)
        if (!out) return
        // Reject a lost race BEFORE pushing undo (same as commitBaked), or ⌘Z
        // would record a phantom step for the no-op map below
        const cur = imageEditsRef.current.find((x) => x.id === target.id)
        if (
          !cur ||
          cur.input !== target.before ||
          (cur.input.kind !== 'insertImage' && cur.input.kind !== 'replaceImage')
        ) {
          return
        }
        pushUndoRef.current()
        setImageEdits((prev) =>
          prev.map((e) =>
            e.id === target.id &&
            e.input === target.before &&
            (e.input.kind === 'insertImage' || e.input.kind === 'replaceImage')
              ? { ...e, opacityBase: prior, input: { ...e.input, image: out } }
              : e,
          ),
        )
        return
      }
      const src = await bakeSourcePng(target)
      const out = src ? await transformPngPixels(src, alpha) : null
      if (out) commitBaked(target, out, undefined, src ?? undefined)
    })()
  }

  /** Crop / remove-background dialog over the selection's fetched pixels */
  const [imageDialog, setImageDialog] = useState<{
    kind: 'crop' | 'cutout'
    target: ImageBakeTarget
    image: string
  } | null>(null)

  const openImageDialog = (kind: 'crop' | 'cutout') => {
    const target = bakeTargetOf(selected)
    if (!target) return
    setSelected(null)
    void bakeSourcePng(target).then((src) => {
      if (src) setImageDialog({ kind, target, image: src })
    })
  }

  /** Replace flow: the bubble button stashes the target, then reuses the insert file input */
  const replaceTargetRef = useRef<ImageBakeTarget | null>(null)
  const startReplaceImage = () => {
    const target = bakeTargetOf(selected)
    if (!target) return
    replaceTargetRef.current = target
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
  // resized/deleted images and without deleted saved annotations, so the original
  // vanishes immediately instead of at save ──

  /** Pages with pending erase ops → image rects to remove + saved annotations to remove */
  const livePreviewRects = useMemo(() => {
    const map = new Map<
      number,
      { rects: [number, number, number, number][]; annots: (SavedMarkupAnnot | SavedNoteAnnot)[] }
    >()
    const jobFor = (pageIndex: number) => {
      let job = map.get(pageIndex)
      if (!job) {
        job = { rects: [], annots: [] }
        map.set(pageIndex, job)
      }
      return job
    }
    for (const e of imageEdits) {
      if (e.input.kind !== 'insertImage') jobFor(e.input.pageIndex).rects.push(e.input.oldRect)
    }
    for (const d of annotDeletes) jobFor(d.annot.pageIndex).annots.push(d.annot)
    return map
  }, [imageEdits, annotDeletes])

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
    for (const [pageIndex, job] of livePreviewRects) {
      const geom = pageGeom(pageIndex)
      const disp = geomDispSize(geom)
      // Union of the erased rects in display coords, padded and clamped to the page
      let x1 = Infinity
      let y1 = Infinity
      let x2 = -Infinity
      let y2 = -Infinity
      for (const r of [...job.rects, ...job.annots.map((a) => a.rect)]) {
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
      const excludedAnnots = job.annots.map((a): AnnotDeleteInput => ({
        pageIndex: a.pageIndex,
        objNum: a.objNum,
        subtype: a.type,
        rect: a.rect,
        ...(a.type === 'note' ? { contents: a.contents } : {}),
      }))
      const annotKey = excludedAnnots
        .map((a) => `${a.objNum}:${a.subtype}:${imageRectKey(a.rect)}`)
        .join(',')
      const key = `${job.rects.map(imageRectKey).join(';')}|${annotKey}|${pxWidth}|${rotIdx}`
      if (livePreviewKeys.current.get(pageIndex) === key) continue
      livePreviewKeys.current.set(pageIndex, key)
      void window.pdfApi
        .pagePreviewPng({
          path: filePath,
          pageIndex,
          excludeRects: job.rects,
          ...(excludedAnnots.length > 0 ? { excludeAnnots: excludedAnnots } : {}),
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

  /** Commit the margin draft card as a pending note; the new thread becomes active */
  const confirmNoteDraft = (text: string) => {
    const target = noteDraft
    setNoteDraft(null)
    if (!target || !text) return
    pushUndo()
    const id = newId()
    setDrawings((prev) => [
      ...prev,
      {
        id,
        input: {
          kind: 'note',
          pageIndex: target.origIdx,
          color: drawColor,
          at: target.at,
          contents: text,
          author: noteAuthor || undefined,
          createdMs: Date.now(),
          localId: id,
        },
      },
    ])
    setActiveNote({ origIdx: target.origIdx, rootKey: pendingNoteKey(id) })
  }

  // ── Note comment threads (saved Text annots + pending notes, WPS-style replies) ──

  /** Pending note drawings on a page (typed extraction; replies included) */
  const pendingNotesOn = (origIdx: number): { id: string; input: NoteInput }[] =>
    drawings.flatMap((d) =>
      d.input.kind === 'note' && d.input.pageIndex === origIdx
        ? [{ id: d.id, input: d.input }]
        : [],
    )

  /** Threads for a page, with notes pending deletion filtered out and pending content
      edits overlaid. Only `item.contents` is overlaid — `item.saved` keeps the on-disk
      text, which replies and the edits themselves need for identity matching at save. */
  const noteThreadsOn = (origIdx: number): NoteThreadItem[] => {
    const pendingDeleted = new Set(annotDeletes.map((d) => d.annot.objNum))
    const saved = (savedNotes.get(origIdx) ?? []).filter((a) => !pendingDeleted.has(a.objNum))
    const roots = buildNoteThreads(saved, pendingNotesOn(origIdx))
    const edited = new Map(
      noteEdits.filter((e) => e.annot.pageIndex === origIdx).map((e) => [e.annot.objNum, e]),
    )
    if (edited.size > 0) {
      for (const root of roots) {
        for (const { item } of flattenThread(root)) {
          const e = item.saved ? edited.get(item.saved.objNum) : undefined
          if (e) item.contents = e.contents
        }
      }
    }
    return roots
  }

  /** Root pins DrawLayer renders for saved threads (pending roots render from drawings) */
  const savedNotePins = (origIdx: number): SavedNotePin[] =>
    noteThreadsOn(origIdx).flatMap((root) =>
      root.saved
        ? [{ key: root.key, at: root.at, color: root.color, contents: root.contents }]
        : [],
    )

  /** Append a reply to a thread's root (flat, WPS-style threads: /IRT → root) */
  const replyToNote = (origIdx: number, root: NoteThreadItem, text: string) => {
    pushUndo()
    const id = newId()
    setDrawings((prev) => [
      ...prev,
      {
        id,
        input: {
          kind: 'note',
          pageIndex: origIdx,
          color: root.color ?? drawColor,
          at: root.at,
          contents: text,
          author: noteAuthor || undefined,
          createdMs: Date.now(),
          localId: id,
          ...(root.saved
            ? {
                replyToSaved: {
                  objNum: root.saved.objNum,
                  rect: root.saved.rect,
                  contents: root.saved.contents,
                },
              }
            : root.pendingId !== null
              ? { replyToLocalId: root.pendingId }
              : {}),
        },
      },
    ])
  }

  /** Next drawings/noteEdits after rewriting one comment's text: pending notes mutate
      the drawing in place; saved notes queue an in-place /Contents edit (keyed by
      note, later edits replace it). Null when nothing changes. */
  const appliedNoteEdit = (
    item: NoteThreadItem,
    trimmed: string,
  ): { drawings: LocalDrawing[]; noteEdits: LocalNoteEdit[] } | null => {
    if (!trimmed || trimmed === item.contents) return null
    if (item.pendingId !== null) {
      const id = item.pendingId
      return {
        drawings: drawings.map((d) =>
          d.id === id && d.input.kind === 'note'
            ? { ...d, input: { ...d.input, contents: trimmed } }
            : d,
        ),
        noteEdits,
      }
    }
    if (item.saved) {
      const annot = item.saved
      // Drop the entry only when the file provably holds this exact text. A save in
      // flight for this note is unconfirmed — keep the entry: if the write lands, the
      // post-save subtraction drops entries matching it; if it fails, the entry still
      // targets the unchanged on-disk text.
      const unconfirmed = inFlightNoteWritesRef.current.has(annot.objNum)
      return {
        drawings,
        noteEdits: [
          ...noteEdits.filter(
            (e) => e.annot.objNum !== annot.objNum || e.annot.pageIndex !== annot.pageIndex,
          ),
          ...(trimmed === annot.contents && !unconfirmed
            ? []
            : [{ id: newId(), annot, contents: trimmed }]),
        ],
      }
    }
    return null
  }

  /** Confirm the comment-edit box (OK button / Cmd+Enter) */
  const editNoteItem = (item: NoteThreadItem, text: string) => {
    const next = appliedNoteEdit(item, text.trim())
    if (!next) return
    pushUndo()
    setDrawings(next.drawings)
    setNoteEdits(next.noteEdits)
  }

  /** Fold an open comment-edit box into pending state (cf. commitTextDraft: a save can
      land mid-typing and the post-save reload tears the edit box down, dropping the
      typed text). Returns the arrays the save must write — the setState calls here
      won't be visible to the caller's closure. */
  const commitNoteEdit = (): { drawings: LocalDrawing[]; noteEdits: LocalNoteEdit[] } => {
    const unchanged = { drawings, noteEdits }
    const draft = noteEditDraft
    if (!draft) return unchanged
    setNoteEditDraft(null)
    const item = noteThreadsOn(draft.origIdx)
      .flatMap((root) => flattenThread(root))
      .map(({ item: it }) => it)
      .find((it) => it.key === draft.itemKey)
    const next = item ? appliedNoteEdit(item, draft.text.trim()) : null
    if (!next) return unchanged
    pushUndo()
    setDrawings(next.drawings)
    setNoteEdits(next.noteEdits)
    return next
  }

  /** Delete a comment and everything under it (saved → pending annotDeletes; pending → dropped) */
  const deleteNoteItem = (item: NoteThreadItem) => {
    pushUndo()
    const { saved, pendingIds } = threadSubtree(item)
    if (saved.length > 0) {
      setAnnotDeletes((prev) => [...prev, ...saved.map((annot) => ({ id: newId(), annot }))])
      // A pending content edit of a deleted note has nothing to apply to anymore
      const gone = new Set(saved.map((a) => a.objNum))
      setNoteEdits((prev) => prev.filter((e) => !gone.has(e.annot.objNum)))
    }
    if (pendingIds.length > 0) {
      const drop = new Set(pendingIds)
      setDrawings((prev) => prev.filter((d) => !drop.has(d.id)))
    }
    if (activeNote && item.key === activeNote.rootKey) setActiveNote(null)
    // "deleted · undo" toast — a mistaken trash tap stays reversible before autosave
    if (toastTimerRef.current !== null) window.clearTimeout(toastTimerRef.current)
    setDeletedInsertedText(false)
    setDeleteToast(true)
    toastTimerRef.current = window.setTimeout(() => setDeleteToast(false), 5000)
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

  const insertBlankPage = (afterOrigIdx: number) => {
    // The flush writes the on-screen order into the file, so the neighbor's
    // post-save index is its visible position, not its pre-save original index
    const afterVis = visList.indexOf(afterOrigIdx)
    return flushThen(async () => {
      const result = await window.pdfApi.insertBlankPage({
        path: filePath,
        afterPageIndex: afterVis,
      })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      await loadDoc(filePath, doc)
    })
  }

  const openSplitDlg = () => {
    setSplitInput('1')
    setSplitInvalid(false)
    setSplitDlg(true)
  }

  /** Split dialog confirm: every N pages becomes its own file (at least two output files) */
  const confirmSplit = () => {
    const n = Number(splitInput.trim())
    if (!Number.isInteger(n) || n < 1 || n >= pageCount) {
      setSplitInvalid(true)
      return
    }
    setSplitDlg(false)
    void flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.splitPdf({ path: filePath, chunkSize: n, baseName: base })
      if (!result.ok) opFailed(result.error)
    })
  }

  const mergePdf = () =>
    flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.mergePdf({
        path: filePath,
        suggestedName: `${base}-merged.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })

  const openMergePagesDlg = () => {
    setMergeCount('2')
    setMergeCountInvalid(false)
    setMergeDirection('vertical')
    setMergeSeparator(false)
    setMergePagesDlg(true)
  }

  /** N-up "merge pages" with WPS-style options; output written to a new file */
  const confirmMergePages = () => {
    const n = Number(mergeCount.trim())
    if (!Number.isInteger(n) || n < 2 || n > 16) {
      setMergeCountInvalid(true)
      return
    }
    setMergePagesDlg(false)
    void flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.mergePages({
        path: filePath,
        perSheet: n,
        direction: mergeDirection,
        separator: mergeSeparator,
        suggestedName: `${base}-${n}in1.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })
  }

  const openReplaceDlg = () => {
    setReplaceInput(String(currentPage))
    setReplaceInvalid(false)
    setReplaceDlg(true)
  }

  /** Replace dialog confirm: visible page-number ranges → original indices; main picks the source PDF */
  const confirmReplace = () => {
    const pages = parsePageRanges(replaceInput, pageCount)
    if (!pages) {
      setReplaceInvalid(true)
      return
    }
    setReplaceDlg(false)
    void flushThen(async () => {
      // The flush writes the on-screen order into the file: visible page n is
      // page n-1 of the saved file (with no unsaved changes visList is identity)
      const result = await window.pdfApi.replacePages({
        path: filePath,
        pages: pages.map((n) => n - 1),
      })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      if (!('canceled' in result)) await loadDoc(filePath, doc)
    })
  }

  /** Resize all pages to the picked paper size (points, portrait) */
  const applyPageSize = (width: number, height: number) => {
    setPageSizeDlg(false)
    void flushThen(async () => {
      const result = await window.pdfApi.setPageSize({ path: filePath, width, height })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      await loadDoc(filePath, doc)
    })
  }

  /** Split every page into a perPage grid, saved as a new file */
  const runSplitPages = (perPage: 2 | 4 | 9) => {
    setSplitPagesDlg(false)
    void flushThen(async () => {
      const base = fileName.replace(/\.pdf$/i, '')
      const result = await window.pdfApi.splitPages({
        path: filePath,
        perPage,
        suggestedName: `${base}-split.pdf`,
      })
      if (!result.ok) opFailed(result.error)
    })
  }

  /** Render the page as displayed (rotation included) for the crop dialog */
  const openPageCrop = async (origIdx: number) => {
    if (!doc) return
    try {
      const page = await doc.getPage(origIdx + 1)
      const rotation = (((page.rotate + rotDelta(origIdx)) % 360) + 360) % 360
      const base = page.getViewport({ scale: 1, rotation })
      const k = Math.min(1600 / base.width, 1600 / base.height, 4)
      const viewport = page.getViewport({ scale: k, rotation })
      const canvas = document.createElement('canvas')
      canvas.width = Math.floor(viewport.width)
      canvas.height = Math.floor(viewport.height)
      await page.render({ canvas, viewport }).promise
      const png = canvas.toDataURL('image/png').split(',')[1]
      canvas.width = 0
      if (!png) return
      setCropAllPages(false)
      setPageCropDlg({ png, origIdx })
    } catch (err) {
      opFailed(err instanceof Error ? err.message : String(err))
    }
  }

  /** Crop confirm: fractions of the displayed page → CropBox shrink on one page or all */
  const confirmPageCrop = (crop: { l: number; t: number; r: number; b: number }) => {
    const dlg = pageCropDlg
    if (!dlg) return
    setPageCropDlg(null)
    // Full-frame selection is a no-op
    if (crop.l <= 0 && crop.t <= 0 && crop.r >= 1 && crop.b >= 1) return
    // The flush writes the on-screen order into the file: post-save indices are
    // the visible positions, not the pre-save original indices
    const pages = cropAllPages ? visList.map((_, i) => i) : [visList.indexOf(dlg.origIdx)]
    if (pages[0]! < 0) return
    void flushThen(async () => {
      const result = await window.pdfApi.cropPages({ path: filePath, pages, rect: crop })
      if (!result.ok) {
        opFailed(result.error)
        return
      }
      await loadDoc(filePath, doc)
    })
  }

  /** Print: save first (markups/forms/page ops all into the file), then reload from the file to render, avoiding a destroyed old doc */
  // Synchronous re-entry guard: the menu accelerator and the renderer's own ⌘P can both
  // fire, and the `printing` state only updates after the async flush has started
  const printBusyRef = useRef(false)
  const printDoc = async () => {
    if (printBusyRef.current) return
    printBusyRef.current = true
    try {
      await flushThen(async () => {
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
    } finally {
      printBusyRef.current = false
    }
  }

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

  // ── unified popover dismissal: a press outside the guard roots, a window
  // blur, or a press on the shell tab strip closes each popover ──

  /** Guard roots (popover panel + its trigger) for the dismissal hooks below */
  const thumbMenuRef = useRef<HTMLDivElement | null>(null)
  const drawColorWrapRef = useRef<HTMLDivElement | null>(null)
  const draftColorWrapRef = useRef<HTMLSpanElement | null>(null)
  const highlightWrapRef = useRef<HTMLDivElement | null>(null)
  const opacityBtnRef = useRef<HTMLButtonElement | null>(null)
  const opacityMenuRef = useRef<HTMLDivElement | null>(null)
  const staticColorFieldRef = useRef<HTMLDivElement | null>(null)

  // Thumbnail context menu
  useDismissablePopover(thumbMenu != null, () => setThumbMenu(null), {
    inside: () => [thumbMenuRef.current],
  })

  // Draw-color palette (Annotate tab). Guarded by its own wrap ref — the
  // generic `.rb-drop-wrap` class is shared with the highlight split button,
  // which used to keep this popover open when that button was pressed.
  useDismissablePopover(colorOpen, () => setColorOpen(false), {
    inside: () => [drawColorWrapRef.current],
  })

  // Text-edit color popover follows its draft: closes with the draft
  useEffect(() => {
    if (!textDraft) setDraftColorOpen(false)
  }, [textDraft])
  useDismissablePopover(draftColorOpen, () => setDraftColorOpen(false), {
    inside: () => [draftColorWrapRef.current],
  })

  // Highlight split-button color popover
  useDismissablePopover(highlightColorOpen, () => setHighlightColorOpen(false), {
    inside: () => [highlightWrapRef.current],
  })

  // Transparency presets fold-out inside the image selection popup
  useDismissablePopover(opacityMenu, () => setOpacityMenu(false), {
    inside: () => [opacityBtnRef.current, opacityMenuRef.current],
  })

  // Color popover inside the static-text dialog (the dialog itself stays open)
  useDismissablePopover(staticTextColorOpen, () => setStaticTextColorOpen(false), {
    inside: () => [staticColorFieldRef.current],
  })

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

  // Shell menu Print → same flow as the ribbon button / ⌘P
  useEffect(() => {
    return window.pdfApi.onPrintRequest(() => void printDoc())
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
        else if (pendingTextInsert) setPendingTextInsert(null)
        else if (imagePick) setImagePick(null)
        else if (editTextMode) setEditTextMode(false)
        else if (editImageMode) setEditImageMode(false)
        else if (pendingSign) setPendingSign(null)
        else if (noteDraft) setNoteDraft(null)
        else if (activeNote) setActiveNote(null)
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
        <div
          ref={highlightWrapRef}
          className="rb-drop-wrap rb-highlight-drop-wrap rb-highlight-split"
        >
          <button
            className="rb-big rb-highlight-main"
            disabled={readOnly}
            data-tip={t('highlight')}
            onClick={() => applyMarkup('highlight')}
          >
            <span className="rb-big-icon">
              <span className="rb-big-icon-colored">
                <IconHighlight />
                <span className="rb-color-bar" style={{ background: cssRgb(highlightColor) }} />
              </span>
            </span>
            {t('highlight')}
          </button>
          <button
            className={`rb-highlight-caret${highlightColorOpen ? ' active' : ''}`}
            disabled={readOnly}
            aria-label={t('drawColor')}
            data-tip={t('drawColor')}
            onClick={() => setHighlightColorOpen((open) => !open)}
          >
            <RbCaret />
          </button>
          {highlightColorOpen && (
            <ColorPickerPopover
              className="rb-drop"
              value={rgbToHex(highlightColor)}
              onPick={(hex) => setHighlightColor(hexToRgb(hex))}
              onClose={() => setHighlightColorOpen(false)}
            />
          )}
        </div>
        <button
          className="rb-big"
          disabled={readOnly}
          data-tip={t('underline')}
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
          data-tip={t('strikeout')}
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
            <button
              className="rb-icon"
              data-tip={t('zoomOut')}
              aria-label={t('zoomOut')}
              onClick={zoomOut}
            >
              −
            </button>
            <span className="tb-zoom">{Math.round(scale * 100)}%</span>
            <button
              className="rb-icon"
              data-tip={t('zoomIn')}
              aria-label={t('zoomIn')}
              onClick={zoomIn}
            >
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
      data-tip={`${t('search')} (${platformShortcuts('⌘F')})`}
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
      data-tip={t('editTextHint')}
      onClick={() => {
        setTextDraft(null)
        setPendingTextInsert(null)
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

  const activeFormWidget = activeFormIndex >= 0 ? formWidgets[activeFormIndex]! : null
  const formWidgetSigned = (widget: FormWidget): boolean =>
    widget.signed || signedFormWidgetIds.has(widget.id)
  const firstSignatureWidget =
    formWidgets.find((widget) => widget.kind === 'signature' && !formWidgetSigned(widget)) ?? null

  const openSignatureDialog = (target: FormWidget | null) => {
    setEditTextMode(false)
    setTextDraft(null)
    setPendingTextInsert(null)
    setDrawTool(null)
    setImagePick(null)
    setEditImageMode(false)
    setPendingSign(null)
    setSignatureTarget(target)
    if (target) focusFormWidget(target)
    setSignDlg(true)
  }

  const completeForm = () => {
    if (missingRequiredFields.length > 0) {
      showNotice(t('formMissingRequired', { count: missingRequiredFields.length }))
      const firstMissing = formWidgets.find(
        (widget) => widget.fieldName === missingRequiredFields[0]!.name,
      )
      if (firstMissing) focusFormWidget(firstMissing)
      return
    }
    showNotice(t('formCompleteDone'))
    setRibbonTab('home')
  }

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
          data-tip={spread === 2 ? t('singlePage') : t('twoPage')}
          onClick={() => setSpread((v) => (v === 1 ? 2 : 1))}
        >
          <span className="rb-big-icon">{spread === 2 ? <IconSinglePage /> : <IconSpread />}</span>
          {spread === 2 ? t('singlePage') : t('twoPage')}
        </button>
        <button
          className={`rb-big${nightMode ? ' active' : ''}`}
          data-tip={t('nightMode')}
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
            data-tip={`${t('save')} (${platformShortcuts('⌘S')})`}
            aria-label={t('save')}
            disabled={!dirty || saveState === 'saving'}
            onClick={() => void save()}
          >
            <IconSave />
          </button>
          <button
            className="qa-btn"
            data-tip={`${t('undo')} (${platformShortcuts('⌘Z')})`}
            aria-label={`${t('undo')} (${platformShortcuts('⌘Z')})`}
            disabled={undoStack.length === 0}
            onClick={undo}
          >
            <IconUndo />
          </button>
          <button
            className="qa-btn"
            data-tip={`${t('redo')} (${platformShortcuts('⇧⌘Z')})`}
            aria-label={`${t('redo')} (${platformShortcuts('⇧⌘Z')})`}
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
          {!readOnly && (
            <button
              className={`ribbon-tab ribbon-tab-context${ribbonTab === 'fillForm' ? ' active' : ''}`}
              onClick={() => setRibbonTab('fillForm')}
            >
              {t('ribbonTabFillForm')}
            </button>
          )}
          <span className="ribbon-tabs-spacer" />
          {readOnly && <span className="tb-readonly">{t('roEncrypted')}</span>}
          {/* The file on disk is only touched by an explicit save until then. */}
          {saveState === 'saving' ? (
            <span className="tb-save-pending">{t('saving')}</span>
          ) : (
            dirty &&
            saveState !== 'error' && <span className="tb-save-pending">{t('unsaved')}</span>
          )}
          {saveState === 'error' && (
            <span className="tb-save-error" data-tip={saveError}>
              {t('saveFailed')}
            </span>
          )}
          {saveState === 'saved' && <span className="tb-save-ok">{t('savedOk')}</span>}
          {formHasXfa && (
            <span className="tb-form-warning" data-tip={t('formXfaWarning')}>
              XFA
            </span>
          )}
        </div>
        <div className="ribbon-body">
          {ribbonTab === 'home' && (
            <>
              {/* ---- Genspark AI (first slot: entry + one-click AI actions, docs parity) ---- */}
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big ai-entry${aiCollapsed ? '' : ' active'}`}
                    data-tip={t('aiOpenAssistant')}
                    onClick={() => setAiCollapsed((v) => !v)}
                  >
                    <span className="rb-big-icon">
                      <GensparkMark size={26} />
                    </span>
                    <span>Genspark AI</span>
                  </button>
                  <button
                    className="rb-big ai-entry"
                    data-tip={t('aiSummarizeBtn')}
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
                    data-tip={t('aiKeyPointsBtn')}
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
                    data-tip={`${t('print')} (${platformShortcuts('⌘P')})`}
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
                    data-tip={t('exportImagesAll')}
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
                    data-tip={t('propsTitle')}
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
                      data-tip={t(key)}
                      onClick={() => {
                        setEditTextMode(false)
                        setTextDraft(null)
                        setPendingTextInsert(null)
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
                    data-tip={t('signTitle')}
                    onClick={() => {
                      if (pendingSign) setPendingSign(null)
                      else openSignatureDialog(null)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconSign />
                    </span>
                    {t('sign')}
                  </button>
                  <div ref={drawColorWrapRef} className="rb-drop-wrap">
                    <button
                      className={`rb-big${colorOpen ? ' active' : ''}`}
                      disabled={readOnly}
                      data-tip={t('drawColor')}
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
                      <ColorPickerPopover
                        className="rb-drop"
                        value={rgbToHex(drawColor)}
                        onPick={(hex) => setDrawColor(hexToRgb(hex))}
                        onClose={() => setColorOpen(false)}
                      />
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
                    className={`rb-big${pendingTextInsert ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('insertTextHint')}
                    onClick={() => {
                      if (pendingTextInsert) {
                        setPendingTextInsert(null)
                        return
                      }
                      setEditTextMode(false)
                      setTextDraft(null)
                      setDrawTool(null)
                      setPendingSign(null)
                      setSignatureTarget(null)
                      setImagePick(null)
                      setPendingStaticFill(null)
                      setEditImageMode(false)
                      setTextInsertEditId(null)
                      setStaticTextPurpose('insert')
                      setStaticText('')
                      setStaticTextDialog(true)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFormText />
                    </span>
                    {t('insertText')}
                  </button>
                  <button
                    className={`rb-big${imagePick && !pendingStaticFill ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('insertImageHint')}
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
                    data-tip={t('editImageHint')}
                    onClick={() => {
                      setEditTextMode(false)
                      setTextDraft(null)
                      setDrawTool(null)
                      setPendingSign(null)
                      setImagePick(null)
                      setPendingTextInsert(null)
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
                    data-tip={t('stampTitle')}
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
          {ribbonTab === 'fillForm' && (
            <>
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className={`rb-big${pendingStaticFill === 'text' ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('formAddTextHint')}
                    onClick={() => {
                      setImagePick(null)
                      setPendingTextInsert(null)
                      setStaticTextEditTarget(null)
                      setStaticTextPurpose('form')
                      setStaticText('')
                      setStaticTextDialog(true)
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconFormText />
                    </span>
                    {t('formAddText')}
                  </button>
                  <button
                    className={`rb-big${pendingStaticFill === 'check' ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('formAddCheckHint')}
                    onClick={() => startStaticFormMark('check')}
                  >
                    <span className="rb-big-icon">
                      <IconFormCheck />
                    </span>
                    {t('formAddCheck')}
                  </button>
                  <button
                    className={`rb-big${pendingStaticFill === 'cross' ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('formAddCrossHint')}
                    onClick={() => startStaticFormMark('cross')}
                  >
                    <span className="rb-big-icon">
                      <IconFormCross />
                    </span>
                    {t('formAddCross')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              {hasFillableForm && (
                <>
                  <div className="ribbon-group">
                    <div className="ribbon-group-items">
                      <button
                        className="rb-big"
                        data-tip={t('formPreviousField')}
                        onClick={() => stepFormWidget(-1)}
                      >
                        <span className="rb-big-icon">
                          <IconPreviousField />
                        </span>
                        {t('formPreviousField')}
                      </button>
                      <button
                        className="rb-big"
                        data-tip={t('formNextField')}
                        onClick={() => stepFormWidget(1)}
                      >
                        <span className="rb-big-icon">
                          <IconNextField />
                        </span>
                        {t('formNextField')}
                      </button>
                      <span className="form-ribbon-progress">
                        {t('formFieldProgress', {
                          current: activeFormIndex >= 0 ? activeFormIndex + 1 : 0,
                          total: formWidgets.length,
                        })}
                      </span>
                    </div>
                  </div>
                  <div className="ribbon-sep" />
                </>
              )}
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    data-tip={t('signTitle')}
                    onClick={() =>
                      openSignatureDialog(
                        activeFormWidget?.kind === 'signature' &&
                          !formWidgetSigned(activeFormWidget)
                          ? activeFormWidget
                          : firstSignatureWidget,
                      )
                    }
                  >
                    <span className="rb-big-icon">
                      <IconSign />
                    </span>
                    {t('sign')}
                  </button>
                  <button
                    className={`rb-big${imagePick && !pendingStaticFill ? ' active' : ''}`}
                    disabled={readOnly}
                    data-tip={t('insertImageHint')}
                    onClick={() => (imagePick ? setImagePick(null) : pickInsertImage())}
                  >
                    <span className="rb-big-icon">
                      <IconInsertImage />
                    </span>
                    {t('insertImage')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button className="rb-big" onClick={completeForm}>
                    <span className="rb-big-icon">
                      <IconCompleteForm />
                    </span>
                    {t('formComplete')}
                  </button>
                </div>
              </div>
            </>
          )}
          {ribbonTab === 'page' && (
            <>
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
                    disabled={pageCount === 0 || readOnly}
                    onClick={() => rotateAllPages(90)}
                  >
                    <span className="rb-big-icon">
                      <IconRotateAll />
                    </span>
                    {t('rotateAllPages')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
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
                  <button
                    className="rb-big"
                    disabled={readOnly}
                    onClick={() => void insertBlankPage(curOrigIdx)}
                  >
                    <span className="rb-big-icon">
                      <IconInsertBlank />
                    </span>
                    {t('insertBlankPage')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={openReplaceDlg}
                  >
                    <span className="rb-big-icon">
                      <IconReplacePages />
                    </span>
                    {t('replacePages')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={curOrigIdx < 0 || readOnly}
                    onClick={() => void openPageCrop(curOrigIdx)}
                  >
                    <span className="rb-big-icon">
                      <IconCropPages />
                    </span>
                    {t('cropPages')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={pageCount === 0 || readOnly}
                    onClick={() => setPageSizeDlg(true)}
                  >
                    <span className="rb-big-icon">
                      <IconPageSize />
                    </span>
                    {t('pageSize')}
                  </button>
                </div>
              </div>
              <div className="ribbon-sep" />
              <div className="ribbon-group">
                <div className="ribbon-group-items">
                  <button
                    className="rb-big"
                    disabled={pageCount <= 1 || readOnly}
                    onClick={reversePages}
                  >
                    <span className="rb-big-icon">
                      <IconReverse />
                    </span>
                    {t('reversePages')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={pageCount <= 1 || readOnly}
                    onClick={openSplitDlg}
                  >
                    <span className="rb-big-icon">
                      <IconSplitPdf />
                    </span>
                    {t('splitPdf')}
                  </button>
                  <button className="rb-big" disabled={readOnly} onClick={() => void mergePdf()}>
                    <span className="rb-big-icon">
                      <IconMergePdf />
                    </span>
                    {t('mergePdf')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={pageCount <= 1 || readOnly}
                    onClick={openMergePagesDlg}
                  >
                    <span className="rb-big-icon">
                      <IconMergePages />
                    </span>
                    {t('mergePages')}
                  </button>
                  <button
                    className="rb-big"
                    disabled={pageCount === 0 || readOnly}
                    onClick={() => setSplitPagesDlg(true)}
                  >
                    <span className="rb-big-icon">
                      <IconSplitPages />
                    </span>
                    {t('splitPages')}
                  </button>
                </div>
              </div>
            </>
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
              data-tip={t('aiOpenAssistant')}
              aria-label={t('aiOpenAssistant')}
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
                        {/* Pending erase ops patch the thumb too, so it tracks the canvas */}
                        {livePreview.has(origIdx) &&
                          (() => {
                            const lp = livePreview.get(origIdx)!
                            return (
                              <img
                                className="pdf-thumb-livepreview"
                                src={`data:image/png;base64,${lp.png}`}
                                alt=""
                                style={{
                                  left: `${(lp.clip.x / size.width) * 100}%`,
                                  top: `${(lp.clip.y / size.height) * 100}%`,
                                  width: `${(lp.clip.width / size.width) * 100}%`,
                                  height: `${(lp.clip.height / size.height) * 100}%`,
                                }}
                              />
                            )
                          })()}
                        {/* Pending edits mirror (markups/drawings/text/images/stamps) */}
                        {visibleThumbs.has(v) &&
                          (() => {
                            const mk = markups.filter((m) => m.pageIndex === origIdx)
                            const dw = drawings.filter((d) => d.input.pageIndex === origIdx)
                            const tes = textEdits.filter((te) => te.input.pageIndex === origIdx)
                            const tis = textInserts.filter((ti) => ti.input.pageIndex === origIdx)
                            const ies = imageEdits.filter((ie) => ie.input.pageIndex === origIdx)
                            const sts = stampPreview.get(origIdx) ?? []
                            if (
                              !mk.length &&
                              !dw.length &&
                              !tes.length &&
                              !tis.length &&
                              !ies.length &&
                              !sts.length
                            )
                              return null
                            return (
                              <ThumbPendingOverlay
                                geom={pageGeom(origIdx)}
                                k={thumbRasterW / size.width}
                                markups={mk}
                                drawings={dw}
                                textEdits={tes}
                                textInserts={tis}
                                imageEdits={ies}
                                stamps={sts}
                              />
                            )
                          })()}
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
                // (markup overlays are pointer-transparent; when hit they stopPropagation
                // in handlePageClick before this runs)
                if (
                  !(e.target as Element).closest?.(
                    '.pdf-draw-shape, .pdf-note-pin, .pdf-stamp-preview, .pdf-textedit-preview, .pdf-textinsert-preview, .pdf-textedit-input, .pdf-imgedit-layer, .pdf-imgedit-under, .pdf-del-popup',
                  )
                )
                  setSelected(null)
                if (!(e.target as Element).closest?.('.pdf-note-pin, .pdf-note-margin')) {
                  setActiveNote(null)
                  // With the note tool armed a click on a page (the draw layer) is a
                  // placement, not a dismissal; clicking anywhere else discards the draft
                  if (drawTool !== 'note' || !(e.target as Element).closest?.('.pdf-draw-layer'))
                    setNoteDraft(null)
                }
              }}
            >
              {rows.map((row, r) => {
                const rowVisible = visibleRows.has(r)
                // Comments margin geometry: pin anchors relative to the margin column,
                // which sits one flex gap right of the row's pages
                const rowWidths = row.map((i) => Math.floor(dispSize(i).width * scale))
                const rowW = rowWidths.reduce((a, b) => a + b, 0) + PAGE_GAP * (row.length - 1)
                const pageOffset = (i: number) =>
                  rowWidths.slice(0, i).reduce((a, b) => a + b, 0) + PAGE_GAP * i
                const pinAnchor = (
                  pageIdxInRow: number,
                  geom: PageGeom,
                  at: [number, number],
                ): [number, number] => {
                  const [vx, vy] = pdfToView(geom, at[0], at[1])
                  return [
                    pageOffset(pageIdxInRow) + vx * scale + 10 - (rowW + PAGE_GAP),
                    vy * scale - 10,
                  ]
                }
                const marginThreads: NoteMarginThread[] =
                  noteMarginOn && rowVisible
                    ? row.flatMap((origIdx, i) =>
                        noteThreadsOn(origIdx).map((root) => {
                          const [pinX, pinY] = pinAnchor(i, pageGeom(origIdx), root.at)
                          return { origIdx, root, pinX, pinY }
                        }),
                      )
                    : []
                const draftIdxInRow = noteDraft ? row.indexOf(noteDraft.origIdx) : -1
                const marginDraft: NoteMarginDraft | null =
                  noteDraft && draftIdxInRow >= 0
                    ? (() => {
                        const [pinX, pinY] = pinAnchor(
                          draftIdxInRow,
                          pageGeom(noteDraft.origIdx),
                          noteDraft.at,
                        )
                        return {
                          // Placement identity (PDF coords, zoom-invariant): a new pin
                          // remounts the draft card so stale text/timestamp never carry over
                          placementKey: `${noteDraft.origIdx}:${noteDraft.at[0]}:${noteDraft.at[1]}`,
                          origIdx: noteDraft.origIdx,
                          pinX,
                          pinY,
                          color: drawColor,
                        }
                      })()
                    : null
                return (
                  <div key={r} ref={setRowRef(r)} data-idx={r} className="pdf-row">
                    {row.map((origIdx) => {
                      const size = dispSize(origIdx)
                      const geom = pageGeom(origIdx)
                      return (
                        <div
                          key={origIdx}
                          className={`pdf-page${editTextMode && !readOnly ? ' pdf-editing-text' : ''}${
                            pendingTextInsert ? ' pdf-inserting-text' : ''
                          }`}
                          style={
                            {
                              width: Math.floor(size.width * scale),
                              height: Math.floor(size.height * scale),
                              '--scale-factor': scale,
                            } as CSSProperties
                          }
                          onClick={(e) => {
                            if (Date.now() - blockDragReleaseAt.current < 400) return
                            if (pendingTextInsert && !readOnly) {
                              const pageBox = e.currentTarget.getBoundingClientRect()
                              placeTextInsert(
                                origIdx,
                                (e.clientX - pageBox.left) / scale,
                                (e.clientY - pageBox.top) / scale,
                              )
                            } else if (editTextMode && !readOnly) startTextEdit(origIdx, e)
                            else handlePageClick(origIdx, e)
                          }}
                          onMouseMove={(e) => {
                            if (pendingTextInsert && !readOnly) {
                              const pageBox = e.currentTarget.getBoundingClientRect()
                              setTextInsertPointer({
                                pageIndex: origIdx,
                                x: (e.clientX - pageBox.left) / scale,
                                y: (e.clientY - pageBox.top) / scale,
                              })
                            } else if (editTextMode && !readOnly) {
                              // move, not over: leaving the hover box across the static textLayer
                              // background fires no over events; updateLineHover cheaply returns
                              // while the anchor span is unchanged.
                              updateLineHover(origIdx, e)
                            }
                          }}
                          onMouseLeave={() => {
                            setTextInsertPointer((pointer) =>
                              pointer?.pageIndex === origIdx ? null : pointer,
                            )
                            if (editTextMode && !readOnly) {
                              clearLineHover()
                              clearBlockHover()
                            }
                          }}
                        >
                          <PdfPage
                            doc={doc}
                            pageNo={origIdx + 1}
                            scale={scale}
                            rotationDelta={rotDelta(origIdx)}
                            visible={rowVisible}
                            onRenderState={pageRenderState}
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
                              title={
                                pendingStaticFill ? t('formPlaceStaticHint') : t('imagePlaceHint')
                              }
                              onPlace={(vx, vy) => placeImage(origIdx, vx, vy)}
                              placeK={pendingStaticFill ? staticFormFillPlaceK : imagePlaceK}
                            />
                          )}
                          {/* Paragraph boxes (WPS-style): every text block outlined while
                            edit-text mode is on; hovered one highlighted, all dimmed
                            while the floating editor is open. The hovered block grows
                            border grips — dragging them moves the whole block. */}
                          {editTextMode &&
                            !readOnly &&
                            rowVisible &&
                            (pageBlocks.get(origIdx) ?? []).map((b, i) => {
                              const drawRect = blockDrawRect(origIdx, b)
                              const box = pdfRectToCss(geom, drawRect, scale)
                              const hovered =
                                blockHover?.origIdx === origIdx && blockHover.idx === i
                              const dragging = blockDrag?.origIdx === origIdx && blockDrag.idx === i
                              const dragCss: CSSProperties = dragging
                                ? {
                                    transform: `translate(${blockDrag.to[0] - blockDrag.from[0]}px, ${
                                      blockDrag.to[1] - blockDrag.from[1]
                                    }px)`,
                                  }
                                : {}
                              const grip = {
                                onPointerDown: (e: React.PointerEvent<HTMLDivElement>) => {
                                  if (e.button !== 0) return
                                  e.preventDefault()
                                  e.stopPropagation()
                                  e.currentTarget.setPointerCapture(e.pointerId)
                                  setBlockDrag({
                                    origIdx,
                                    idx: i,
                                    from: [e.clientX, e.clientY],
                                    to: [e.clientX, e.clientY],
                                  })
                                },
                                onPointerMove: (e: React.PointerEvent<HTMLDivElement>) => {
                                  setBlockDrag((cur) =>
                                    cur ? { ...cur, to: [e.clientX, e.clientY] } : cur,
                                  )
                                },
                                onPointerUp: (e: React.PointerEvent<HTMLDivElement>) => {
                                  blockDragReleaseAt.current = Date.now()
                                  const dr = blockDrag
                                  setBlockDrag(null)
                                  if (!dr) return
                                  const dxPx = e.clientX - dr.from[0]
                                  const dyPx = e.clientY - dr.from[1]
                                  if (Math.hypot(dxPx, dyPx) < 3) return
                                  const [ax, ay] = viewToPdf(geom, 0, 0)
                                  const [bx, by] = viewToPdf(geom, dxPx / scale, dyPx / scale)
                                  commitBlockMove(origIdx, b, [bx - ax, by - ay])
                                },
                                onPointerCancel: () => setBlockDrag(null),
                              }
                              return (
                                <Fragment key={i}>
                                  <div
                                    className={`pdf-textblock-box${hovered ? ' is-hover' : ''}${
                                      textDraft ? ' is-faded' : ''
                                    }`}
                                    style={{ ...box, ...dragCss }}
                                  />
                                  {(hovered || dragging) && !textDraft && (
                                    <div
                                      className="pdf-textblock-frame"
                                      style={{ ...box, ...dragCss }}
                                      data-tip={t('textBlockMoveHint')}
                                    >
                                      <div className="pdf-textblock-grip grip-t" {...grip} />
                                      <div className="pdf-textblock-grip grip-b" {...grip} />
                                      <div className="pdf-textblock-grip grip-l" {...grip} />
                                      <div className="pdf-textblock-grip grip-r" {...grip} />
                                    </div>
                                  )}
                                  {dragging &&
                                    (() => {
                                      // The ghost previews the ORIGINAL text following the
                                      // pointer; rows a pending edit claims are skipped —
                                      // their styled previews follow the drag instead, so
                                      // the user never sees the pre-edit look mid-drag
                                      const claims = blockClaimEdits(origIdx, b)
                                      const rows = b.lines
                                        .map((l, j) => [l, j] as const)
                                        .filter(
                                          ([l]) =>
                                            !claims.some(
                                              (te) =>
                                                overlapOfSmaller(te.input.rect, l.rect) >= 0.5,
                                            ),
                                        )
                                      if (rows.length === 0) return null
                                      return (
                                        <div
                                          className="pdf-textblock-ghost"
                                          style={{ ...box, ...dragCss }}
                                        >
                                          {rows.map(([l, j]) => {
                                            const lb = pdfRectToCss(
                                              geom,
                                              shiftRect(l.rect, [
                                                drawRect[0] - b.rect[0],
                                                drawRect[1] - b.rect[1],
                                              ]),
                                              scale,
                                            )
                                            const gf = l.font ? docFonts.get(l.font) : undefined
                                            return (
                                              <div
                                                key={j}
                                                className="pdf-textblock-ghost-line"
                                                style={{
                                                  left: lb.left - box.left,
                                                  top: lb.top - box.top,
                                                  height: lb.height,
                                                  fontSize: l.fontSize * scale * 0.92,
                                                  ...(gf ? { fontFamily: gf.css } : {}),
                                                  ...(gf?.weight ? { fontWeight: gf.weight } : {}),
                                                  ...(gf?.italic
                                                    ? { fontStyle: 'italic' as const }
                                                    : {}),
                                                }}
                                              >
                                                {l.text}
                                              </div>
                                            )
                                          })}
                                        </div>
                                      )
                                    })()}
                                </Fragment>
                              )
                            })}
                          {editTextMode && !readOnly && lineHover?.origIdx === origIdx && (
                            <div className="pdf-textline-hover" style={lineHover.box} />
                          )}
                          {rowVisible && (
                            <>
                              {pendingTextInsert && textInsertPointer?.pageIndex === origIdx && (
                                <div
                                  className="pdf-textinsert-placement-preview"
                                  style={{
                                    left: textInsertPointer.x * scale,
                                    top: (textInsertPointer.y - pendingTextInsert.fontSize) * scale,
                                    fontSize: pendingTextInsert.fontSize * scale * 0.92,
                                    lineHeight: pendingTextInsert.lineLeading
                                      ? `${pendingTextInsert.lineLeading * scale}px`
                                      : 1.2,
                                    color: `rgb(${pendingTextInsert.color.join(', ')})`,
                                    whiteSpace: 'pre',
                                    transform:
                                      pendingTextInsert.align === 'center'
                                        ? 'translateX(-50%)'
                                        : pendingTextInsert.align === 'right'
                                          ? 'translateX(-100%)'
                                          : undefined,
                                    textAlign: pendingTextInsert.align ?? 'left',
                                  }}
                                >
                                  {pendingTextInsert.text}
                                </div>
                              )}
                              {textInserts
                                .filter((insert) => insert.input.pageIndex === origIdx)
                                .map((insert) => {
                                  const style = textInsertPreviewStyle(insert, geom, scale)
                                  if (insertDrag?.id === insert.id) {
                                    // Live drag: ride along under the pointer; the align shift
                                    // (translateX) composes additively with the drag offset
                                    style.transform = `translate(${insertDrag.to[0] - insertDrag.from[0]}px, ${
                                      insertDrag.to[1] - insertDrag.from[1]
                                    }px)${style.transform ? ` ${style.transform}` : ''}`
                                  }
                                  return (
                                    <div
                                      key={insert.id}
                                      className={`pdf-textinsert-preview${
                                        selected?.kind === 'textInsert' && selected.id === insert.id
                                          ? ' is-selected'
                                          : ''
                                      }`}
                                      style={style}
                                      onPointerDown={(e) => {
                                        if (e.button !== 0) return
                                        e.preventDefault()
                                        e.stopPropagation()
                                        e.currentTarget.setPointerCapture(e.pointerId)
                                        setInsertDrag({
                                          id: insert.id,
                                          from: [e.clientX, e.clientY],
                                          to: [e.clientX, e.clientY],
                                        })
                                      }}
                                      onPointerMove={(e) => {
                                        setInsertDrag((cur) =>
                                          cur && cur.id === insert.id
                                            ? { ...cur, to: [e.clientX, e.clientY] }
                                            : cur,
                                        )
                                      }}
                                      onPointerUp={(e) => {
                                        const dr = insertDrag
                                        setInsertDrag(null)
                                        if (!dr || dr.id !== insert.id) return
                                        const dxPx = e.clientX - dr.from[0]
                                        const dyPx = e.clientY - dr.from[1]
                                        // Below the threshold it's a plain click: selection runs
                                        if (Math.hypot(dxPx, dyPx) < 3) return
                                        blockDragReleaseAt.current = Date.now()
                                        const [ax, ay] = viewToPdf(geom, 0, 0)
                                        const [bx, by] = viewToPdf(geom, dxPx / scale, dyPx / scale)
                                        pushUndo()
                                        setTextInserts((prev) =>
                                          prev.map((it) =>
                                            it.id === insert.id
                                              ? {
                                                  ...it,
                                                  input: {
                                                    ...it.input,
                                                    origin: [
                                                      it.input.origin[0] + (bx - ax),
                                                      it.input.origin[1] + (by - ay),
                                                    ],
                                                  },
                                                }
                                              : it,
                                          ),
                                        )
                                        setSelected(null)
                                      }}
                                      onPointerCancel={() => setInsertDrag(null)}
                                      onClick={(e) => {
                                        e.stopPropagation()
                                        // The click synthesized from a completed drag must not
                                        // pop the selection at the drop point
                                        if (Date.now() - blockDragReleaseAt.current < 400) return
                                        setSelected({
                                          kind: 'textInsert',
                                          id: insert.id,
                                          ...popupPos(e.clientX, e.clientY),
                                        })
                                      }}
                                      onDoubleClick={(e) => {
                                        e.stopPropagation()
                                        setSelected(null)
                                        setPendingTextInsert(null)
                                        setTextInsertEditId(insert.id)
                                        setStaticTextPurpose('insert')
                                        setStaticText(insert.input.text)
                                        setStaticTextSize(insert.input.fontSize)
                                        setStaticTextColor(rgb255ToHex(insert.input.color))
                                        setStaticTextAlign(insert.input.align ?? 'left')
                                        setStaticTextDialog(true)
                                      }}
                                    >
                                      {insert.input.text}
                                    </div>
                                  )
                                })}
                              {/* Pending text edits: cover the original run and preview the replacement */}
                              {textEdits
                                .filter((te) => te.input.pageIndex === origIdx)
                                .map((te) => {
                                  const { style, coverStyle } = textEditPreviewParts(
                                    te,
                                    geom,
                                    scale,
                                  )
                                  // A border-drag of the block this edit addresses moves the
                                  // styled preview live with the pointer (the drag ghost of
                                  // the ORIGINAL text is suppressed for edited blocks — it
                                  // would show the pre-edit look while dragging); the cover
                                  // stays put, it hides the original ink
                                  const draggedBlock =
                                    blockDrag && blockDrag.origIdx === origIdx
                                      ? pageBlocks.get(origIdx)?.[blockDrag.idx]
                                      : undefined
                                  if (
                                    blockDrag &&
                                    draggedBlock &&
                                    (blockOwnerEdit(origIdx, draggedBlock)?.id === te.id ||
                                      blockClaimEdits(origIdx, draggedBlock).some(
                                        (c) => c.id === te.id,
                                      ))
                                  ) {
                                    style.transform = `translate(${
                                      blockDrag.to[0] - blockDrag.from[0]
                                    }px, ${blockDrag.to[1] - blockDrag.from[1]}px)`
                                    style.pointerEvents = 'none'
                                    // Above the drag ghost (z 5): the ghost's paper
                                    // background would hide the styled preview riding
                                    // along in the block's claimed rows
                                    style.zIndex = 6
                                  }
                                  return (
                                    <Fragment key={te.id}>
                                      {coverStyle && (
                                        <div className="pdf-textedit-cover" style={coverStyle} />
                                      )}
                                      <div
                                        className={`pdf-textedit-preview${
                                          selected?.kind === 'textEdit' && selected.id === te.id
                                            ? ' pdf-textedit-selected'
                                            : ''
                                        }`}
                                        style={style}
                                        data-tip={
                                          editTextMode ? t('editTextHint') : t('removeMarkup')
                                        }
                                        onClick={(e) => {
                                          e.stopPropagation()
                                          if (editTextMode && !readOnly) {
                                            draftSelectedRef.current = false
                                            // A line edit inside a multi-line clustered
                                            // paragraph reopens the whole paragraph with the
                                            // pending change (and its styles) folded in —
                                            // re-clicking edited text must not demote
                                            // paragraph editing to that single line
                                            if (isFoldableLineEdit(te.input)) {
                                              const r = te.input.rect
                                              const cx = (r[0] + r[2]) / 2
                                              const cy = (r[1] + r[3]) / 2
                                              const block = pageBlocks
                                                .get(origIdx)
                                                ?.find(
                                                  (b) =>
                                                    b.lines.length > 1 &&
                                                    cx >= b.rect[0] &&
                                                    cx <= b.rect[2] &&
                                                    cy >= b.rect[1] &&
                                                    cy <= b.rect[3],
                                                )
                                              if (
                                                block &&
                                                foldBlockValue(
                                                  origIdx,
                                                  block,
                                                  joinBlockLines(block.lines.map((l) => l.text)),
                                                )
                                              ) {
                                                startBlockEdit(origIdx, block)
                                                return
                                              }
                                            }
                                            const draft = pendingDraft(te)
                                            // Unified selection model: the reopened draft's
                                            // caret goes where the preview was clicked. The
                                            // preview shows wrapped newText; the click's
                                            // offset in its textContent maps onto the
                                            // draft's logical value like a line-in-block.
                                            {
                                              const host = e.currentTarget
                                              const cr = document.caretRangeFromPoint(
                                                e.clientX,
                                                e.clientY,
                                              )
                                              let pre: [number, number] | null = null
                                              if (cr && host.contains(cr.startContainer)) {
                                                let off = 0
                                                const walk = document.createTreeWalker(
                                                  host,
                                                  NodeFilter.SHOW_TEXT,
                                                )
                                                for (
                                                  let n = walk.nextNode();
                                                  n;
                                                  n = walk.nextNode()
                                                ) {
                                                  if (n === cr.startContainer) {
                                                    pre = mapLineRangeToBlock(
                                                      draft.value,
                                                      host.textContent ?? '',
                                                      off + cr.startOffset,
                                                      off + cr.startOffset,
                                                    )
                                                    break
                                                  }
                                                  off += (n.textContent ?? '').length
                                                }
                                              }
                                              draftPreselectRef.current = pre
                                            }
                                            setTextDraft(draft)
                                          } else {
                                            setSelected({
                                              kind: 'textEdit',
                                              id: te.id,
                                              ...popupPos(e.clientX, e.clientY),
                                            })
                                          }
                                        }}
                                      >
                                        {textEditPreviewContent(te, scale)}
                                      </div>
                                    </Fragment>
                                  )
                                })}
                              {textDraft &&
                                textDraft.origIdx === origIdx &&
                                (() => {
                                  // A reopened moved edit edits at its new position; rect
                                  // itself stays the original (save-time match key + cover)
                                  const box = pdfRectToCss(
                                    geom,
                                    textDraft.moveBy
                                      ? shiftRect(textDraft.rect, textDraft.moveBy)
                                      : textDraft.rect,
                                    scale,
                                  )
                                  const fs = (textDraft.size ?? textDraft.fontSize) * scale * 0.92
                                  const lines = textDraft.value.split('\n')
                                  // Look-alike of the document's face, so the editor
                                  // reads like the page (until the user picks a font)
                                  const seedF = textDraft.font ? undefined : textDraft.seedFont
                                  const draftCss = textDraft.font
                                    ? EDIT_FONT_BY_ID.get(textDraft.font)?.css
                                    : seedF?.css
                                  const bodyFamily = getComputedStyle(document.body).fontFamily
                                  const blk = textDraft.block
                                  const sizePt = textDraft.size ?? textDraft.fontSize
                                  const draftStyle = `${
                                    textDraft.italic || seedF?.italic ? 'italic ' : ''
                                  }${
                                    textDraft.bold
                                      ? 'bold'
                                      : seedF?.weight
                                        ? String(seedF.weight)
                                        : ''
                                  }`.trim()
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
                                  // Selection-level styles: the textarea can't render
                                  // mixed styles, so its text goes transparent and a
                                  // mirror behind the caret shows them (metric-identical
                                  // for colors; face/size/weight overrides may drift a
                                  // little from the caret until the draft commits)
                                  const draftStyles = textDraft.charStyles?.some((c) => c)
                                    ? textDraft.charStyles
                                    : undefined
                                  const longest = Math.max(
                                    ...lines.map((l) => measureTextWidth(l, editorFont)),
                                  )
                                  const pageEdgeCap =
                                    geomDispSize(geom).width * scale - box.left - 8
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
                                        <div ref={textEditBarRef} className="pdf-textedit-bar">
                                          {editFonts.length > 0 && (
                                            <select
                                              className="pdf-textedit-fontsel"
                                              data-tip={t('texteditFont')}
                                              value={textDraft.font ?? ''}
                                              onChange={(e) => {
                                                const id = e.target.value || undefined
                                                applyDraftStyle(
                                                  // "Original font" on a selection clears
                                                  // the override back to the draft level
                                                  () => ({ font: id ?? null }),
                                                  (d) => ({ ...d, font: id }),
                                                  true,
                                                )
                                              }}
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
                                            data-tip={t('watermarkSize')}
                                            value={
                                              textDraft.size ??
                                              Math.round(textDraft.fontSize * 10) / 10
                                            }
                                            onChange={(e) => {
                                              const v = Number(e.target.value)
                                              if (v >= 1) {
                                                applyDraftStyle(
                                                  () => ({ size: v }),
                                                  (d) => ({ ...d, size: v }),
                                                  false,
                                                )
                                              }
                                            }}
                                          />
                                          <button
                                            className={`pdf-textedit-toggle${textDraft.bold ? ' active' : ''}`}
                                            data-tip={t('texteditBold')}
                                            onClick={() => toggleDraftFlag('bold')}
                                          >
                                            B
                                          </button>
                                          <button
                                            className={`pdf-textedit-toggle pdf-textedit-toggle-i${
                                              textDraft.italic ? ' active' : ''
                                            }`}
                                            data-tip={t('texteditItalic')}
                                            onClick={() => toggleDraftFlag('italic')}
                                          >
                                            I
                                          </button>
                                          <span
                                            ref={draftColorWrapRef}
                                            className="pdf-textedit-colorwrap"
                                          >
                                            <button
                                              className="pdf-textedit-swatch"
                                              style={{
                                                background:
                                                  textDraft.color ?? textDraft.seedInk ?? '#000000',
                                              }}
                                              data-tip={t('drawColor')}
                                              aria-label={t('drawColor')}
                                              onClick={() => setDraftColorOpen((v) => !v)}
                                            />
                                            {draftColorOpen && (
                                              <ColorPickerPopover
                                                value={textDraft.color ?? textDraft.seedInk}
                                                onPick={(hex) => applyDraftColor(hex, true)}
                                                onClose={() => setDraftColorOpen(false)}
                                              />
                                            )}
                                          </span>
                                          <button
                                            className="pdf-textedit-toggle pdf-textedit-trash"
                                            data-tip={t('texteditDeleteRun')}
                                            aria-label={t('texteditDeleteRun')}
                                            onClick={deleteDraftRun}
                                          >
                                            <IconTrash />
                                          </button>
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
                                            // Document-content color (user's pick, else the
                                            // document's own ink), not chrome
                                            ...(textDraft.color || textDraft.seedInk
                                              ? { color: textDraft.color ?? textDraft.seedInk }
                                              : {}),
                                            ...(draftStyles
                                              ? {
                                                  color: 'transparent',
                                                  caretColor:
                                                    textDraft.color ??
                                                    textDraft.seedInk ??
                                                    'var(--pdf-textedit-ink)',
                                                }
                                              : {}),
                                            ...(draftCss ? { fontFamily: draftCss } : {}),
                                            ...(textDraft.bold
                                              ? { fontWeight: 700 }
                                              : seedF?.weight
                                                ? { fontWeight: seedF.weight }
                                                : {}),
                                            ...(textDraft.italic || seedF?.italic
                                              ? { fontStyle: 'italic' }
                                              : {}),
                                          }}
                                          value={textDraft.value}
                                          autoFocus
                                          onFocus={(e) => {
                                            if (!draftSelectedRef.current) {
                                              draftSelectedRef.current = true
                                              const pre = draftPreselectRef.current
                                              draftPreselectRef.current = null
                                              const len = e.currentTarget.value.length
                                              if (pre)
                                                e.currentTarget.setSelectionRange(
                                                  Math.min(pre[0], len),
                                                  Math.min(pre[1], len),
                                                )
                                              else e.currentTarget.setSelectionRange(len, len)
                                            }
                                          }}
                                          onChange={(e) => {
                                            const v = e.target.value
                                            setTextDraft((d) =>
                                              d
                                                ? {
                                                    ...d,
                                                    value: v,
                                                    charStyles: d.charStyles?.some((c) => c)
                                                      ? spliceCharColors(d.value, d.charStyles, v)
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
                                            } else if (
                                              (e.key === 'Backspace' || e.key === 'Delete') &&
                                              textDraft.value === ''
                                            ) {
                                              // The editor is already empty: one more delete
                                              // press removes the run itself
                                              e.preventDefault()
                                              deleteDraftRun()
                                            }
                                          }}
                                        />
                                        {draftStyles && (
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
                                              color:
                                                textDraft.color ??
                                                textDraft.seedInk ??
                                                'var(--pdf-textedit-ink)',
                                              ...(draftCss ? { fontFamily: draftCss } : {}),
                                              ...(textDraft.bold
                                                ? { fontWeight: 700 }
                                                : seedF?.weight
                                                  ? { fontWeight: seedF.weight }
                                                  : {}),
                                              ...(textDraft.italic || seedF?.italic
                                                ? { fontStyle: 'italic' }
                                                : {}),
                                            }}
                                          >
                                            {colorSegments(textDraft.value, draftStyles).map(
                                              (seg, i) => {
                                                if (!seg.color)
                                                  return <Fragment key={i}>{seg.text}</Fragment>
                                                const s = decodeStyle(seg.color)
                                                return (
                                                  <span key={i} style={styleSegCss(s, scale)}>
                                                    {seg.text}
                                                  </span>
                                                )
                                              },
                                            )}
                                          </div>
                                        )}
                                      </div>
                                    </>
                                  )
                                })()}
                              {(imageEdits.some((ie) => ie.input.pageIndex === origIdx) ||
                                (ribbonTab === 'fillForm' &&
                                  savedStaticFormFills.some(
                                    (record) => record.pageIndex === origIdx,
                                  )) ||
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
                                    edits={imageEdits.filter(
                                      (ie) => ie.input.pageIndex === origIdx,
                                    )}
                                    existing={[
                                      ...(editImageMode
                                        ? pageImages.filter((ref) => ref.pageIndex === origIdx)
                                        : []),
                                      ...(ribbonTab === 'fillForm'
                                        ? savedStaticFormFills
                                            .filter((record) => record.pageIndex === origIdx)
                                            .map((record): PageImageRef => ({
                                              pageIndex: record.pageIndex,
                                              rect: record.rect,
                                              aboveText: true,
                                            }))
                                        : []),
                                    ].filter(
                                      (ref, index, refs) =>
                                        !claimedImageKeys.has(
                                          `${ref.pageIndex}:${imageRectKey(ref.rect)}`,
                                        ) &&
                                        refs.findIndex(
                                          (candidate) =>
                                            candidate.pageIndex === ref.pageIndex &&
                                            rectsNear(candidate.rect, ref.rect),
                                        ) === index,
                                    )}
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
                                  data-tip={t('removeStamp')}
                                  style={{
                                    ...pdfRectToCss(geom, s.rect, scale),
                                    opacity: s.opacity ?? 1,
                                  }}
                                  onClick={(e) =>
                                    setSelected({
                                      kind: 'stamp',
                                      ...popupPos(e.clientX, e.clientY),
                                    })
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
                              />
                              {/* Selection outline for a saved markup annotation (the markup
                                itself is painted in the canvas raster) */}
                              {selected?.kind === 'savedMarkup' &&
                                selected.annot.pageIndex === origIdx &&
                                selected.annot.quads.map((q, i) => (
                                  <div
                                    key={i}
                                    className="pdf-markup pdf-markup-selected"
                                    style={pdfRectToCss(geom, quadToRect(q), scale)}
                                  />
                                ))}
                              <DrawLayer
                                geom={geom}
                                scale={scale}
                                pageWidth={size.width}
                                pageHeight={size.height}
                                drawings={drawings.filter((d) => d.input.pageIndex === origIdx)}
                                savedNotes={savedNotePins(origIdx)}
                                activeNoteKey={
                                  activeNote?.origIdx === origIdx ? activeNote.rootKey : null
                                }
                                tool={readOnly ? null : drawTool}
                                color={drawColor}
                                strokeWidth={STROKE_WIDTH}
                                selectedId={selected?.kind === 'drawing' ? selected.id : null}
                                selectTitle={t('removeMarkup')}
                                noteOpenTitle={t('noteOpen')}
                                onCommit={(input) => commitDrawing(origIdx, input)}
                                onNoteAt={(at) => {
                                  setActiveNote(null)
                                  setNoteDraft({ origIdx, at })
                                }}
                                onNoteOpen={(key) => {
                                  setNoteDraft(null)
                                  setActiveNote((prev) =>
                                    prev?.origIdx === origIdx && prev.rootKey === key
                                      ? null
                                      : { origIdx, rootKey: key },
                                  )
                                }}
                                onSelect={(id, x, y) =>
                                  setSelected({ kind: 'drawing', id, ...popupPos(x, y) })
                                }
                                onMove={readOnly ? undefined : moveDrawing}
                                onResize={readOnly ? undefined : resizeDrawing}
                              />
                              {/* Ghost pin for the note being typed into the margin draft card */}
                              {noteDraft?.origIdx === origIdx &&
                                (() => {
                                  const [vx, vy] = pdfToView(geom, noteDraft.at[0], noteDraft.at[1])
                                  return (
                                    <div
                                      className="pdf-note-pin pdf-note-pin-ghost"
                                      style={{
                                        left: vx * scale,
                                        top: vy * scale - 20,
                                        background: cssRgb(drawColor),
                                      }}
                                    >
                                      <svg
                                        width="11"
                                        height="11"
                                        viewBox="0 0 16 16"
                                        fill="none"
                                        stroke="#fff"
                                        strokeWidth="1.6"
                                        aria-hidden
                                      >
                                        <path
                                          d="M2.5 3.5h11v8h-6l-3 2.5V11.5h-2z"
                                          strokeLinejoin="round"
                                        />
                                      </svg>
                                    </div>
                                  )
                                })()}
                              <LinkLayer
                                doc={doc}
                                pageNo={origIdx + 1}
                                geom={geom}
                                scale={scale}
                                onGoToDest={(dest) => void goToDest(dest)}
                              />
                              <FormLayer
                                widgets={formCatalog?.byPage.get(origIdx) ?? []}
                                geom={geom}
                                scale={scale}
                                readOnly={readOnly}
                                edits={formEdits}
                                activeWidgetId={activeFormWidgetId}
                                signedWidgetIds={signedFormWidgetIds}
                                signatureLabel={t('formSignField')}
                                registerControl={registerFormControl}
                                onFocus={(widget) => {
                                  setActiveFormWidgetId(widget.id)
                                  setRibbonTab('fillForm')
                                }}
                                onSignature={(widget) => openSignatureDialog(widget)}
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
                    {noteMarginOn && (
                      <NoteMarginColumn
                        width={NOTE_MARGIN_W}
                        threads={marginThreads}
                        draft={marginDraft}
                        activeKey={
                          activeNote && row.includes(activeNote.origIdx) ? activeNote.rootKey : null
                        }
                        author={noteAuthor}
                        lang={lang}
                        readOnly={readOnly}
                        t={t}
                        onActivate={(th) => {
                          setNoteDraft(null)
                          setActiveNote({ origIdx: th.origIdx, rootKey: th.root.key })
                        }}
                        onReply={(th, text) => replyToNote(th.origIdx, th.root, text)}
                        editingKey={noteEditDraft?.itemKey ?? null}
                        editingText={noteEditDraft?.text ?? ''}
                        onEditStart={(th, item) =>
                          setNoteEditDraft({
                            origIdx: th.origIdx,
                            rootKey: th.root.key,
                            itemKey: item.key,
                            text: item.contents,
                          })
                        }
                        onEditChange={(text) =>
                          setNoteEditDraft((prev) => (prev ? { ...prev, text } : prev))
                        }
                        onEditCancel={() => setNoteEditDraft(null)}
                        onEditSubmit={(_th, item) => {
                          if (noteEditDraft) editNoteItem(item, noteEditDraft.text)
                          setNoteEditDraft(null)
                        }}
                        onDeleteItem={(_th, item) => deleteNoteItem(item)}
                        onClose={() => setActiveNote(null)}
                        onDraftConfirm={confirmNoteDraft}
                        onDraftCancel={() => setNoteDraft(null)}
                      />
                    )}
                  </div>
                )
              })}
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
                  data-tip={t('searchPrev')}
                  aria-label={t('searchPrev')}
                  disabled={activeMatches.length === 0}
                  onClick={() => searchStep(-1)}
                >
                  ‹
                </button>
                <button
                  className="rb-icon"
                  data-tip={t('searchNext')}
                  aria-label={t('searchNext')}
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
                  className={activeMarkupTypes.has('highlight') ? 'is-active' : undefined}
                  data-tip={activeMarkupTypes.has('highlight') ? t('removeMarkup') : t('highlight')}
                  aria-label={
                    activeMarkupTypes.has('highlight') ? t('removeMarkup') : t('highlight')
                  }
                  onClick={() => applyMarkup('highlight')}
                >
                  <span
                    className="sel-swatch sel-swatch-hl"
                    style={{ background: cssRgb(highlightColor) }}
                  />
                </button>
                <button
                  type="button"
                  className={activeMarkupTypes.has('underline') ? 'is-active' : undefined}
                  data-tip={activeMarkupTypes.has('underline') ? t('removeMarkup') : t('underline')}
                  aria-label={
                    activeMarkupTypes.has('underline') ? t('removeMarkup') : t('underline')
                  }
                  onClick={() => applyMarkup('underline')}
                >
                  <span className="sel-swatch sel-swatch-ul">U</span>
                </button>
                <button
                  type="button"
                  className={activeMarkupTypes.has('strikeout') ? 'is-active' : undefined}
                  data-tip={activeMarkupTypes.has('strikeout') ? t('removeMarkup') : t('strikeout')}
                  aria-label={
                    activeMarkupTypes.has('strikeout') ? t('removeMarkup') : t('strikeout')
                  }
                  onClick={() => applyMarkup('strikeout')}
                >
                  <span className="sel-swatch sel-swatch-st">S</span>
                </button>
              </div>
            )}
            {selected && (
              <div
                ref={delPopupRef}
                className="pdf-del-popup"
                style={{ left: selected.x, top: selected.y }}
                onMouseDown={(e) => e.preventDefault()}
              >
                {selectedStaticTextTarget() && (
                  <>
                    <button
                      type="button"
                      data-tip={t('formEditText')}
                      aria-label={t('formEditText')}
                      onClick={startEditStaticText}
                    >
                      <IconFormText />
                    </button>
                    <span className="pdf-del-popup-sep" />
                  </>
                )}
                {selectedImageLayer() !== null && (
                  <>
                    <button
                      type="button"
                      data-tip={t('imageRotateCw')}
                      aria-label={t('imageRotateCw')}
                      onClick={() => rotateSelected(1)}
                    >
                      <IconRotateCw />
                    </button>
                    <button
                      type="button"
                      data-tip={t('imageRotateCcw')}
                      aria-label={t('imageRotateCcw')}
                      onClick={() => rotateSelected(-1)}
                    >
                      <IconRotateCcw />
                    </button>
                    <button
                      type="button"
                      data-tip={t('imageFlipH')}
                      aria-label={t('imageFlipH')}
                      onClick={() => flipSelected('h')}
                    >
                      <IconFlipH />
                    </button>
                    <button
                      type="button"
                      data-tip={t('imageFlipV')}
                      aria-label={t('imageFlipV')}
                      onClick={() => flipSelected('v')}
                    >
                      <IconFlipV />
                    </button>
                    <span className="pdf-del-popup-sep" />
                    <button
                      type="button"
                      data-tip={t('imageCrop')}
                      aria-label={t('imageCrop')}
                      onClick={() => openImageDialog('crop')}
                    >
                      <IconCrop />
                    </button>
                    <button
                      type="button"
                      data-tip={t('imageCutout')}
                      aria-label={t('imageCutout')}
                      onClick={() => openImageDialog('cutout')}
                    >
                      <IconCutout />
                    </button>
                    <button
                      ref={opacityBtnRef}
                      type="button"
                      data-tip={t('imageOpacity')}
                      aria-label={t('imageOpacity')}
                      onClick={() => setOpacityMenu((v) => !v)}
                    >
                      <IconOpacity />
                    </button>
                    <button
                      type="button"
                      data-tip={t('imageReplace')}
                      aria-label={t('imageReplace')}
                      onClick={startReplaceImage}
                    >
                      <IconSwapImage />
                    </button>
                    {opacityMenu && (
                      <div ref={opacityMenuRef} className="pdf-opacity-menu">
                        {[0, 15, 30, 50, 65, 80, 95].map((p) => (
                          <button
                            key={p}
                            type="button"
                            onClick={() => {
                              setOpacityMenu(false)
                              applyImageOpacity(p)
                            }}
                          >
                            {p}%
                          </button>
                        ))}
                      </div>
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
                      : selected.kind === 'textInsert'
                        ? 'deleteInsertedText'
                        : 'deleteAnnotation',
                  )}
                </button>
              </div>
            )}
            {deleteToast && (
              <div className="pdf-toast">
                <span>{t(deletedInsertedText ? 'insertedTextDeleted' : 'annotationDeleted')}</span>
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
              <div
                ref={thumbMenuRef}
                className="thumb-menu file-menu"
                style={{ left: thumbMenu.x, top: thumbMenu.y }}
              >
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
                <button
                  onClick={() => {
                    setThumbMenu(null)
                    void insertBlankPage(menuOrig)
                  }}
                >
                  {t('insertBlankPage')}
                </button>
              </div>
            )}
            {stampDlg && (
              <StampDialog t={t} onCancel={() => setStampDlg(false)} onApply={applyStamps} />
            )}
            {imageDialog?.kind === 'crop' && (
              <CropDialog
                t={t}
                image={imageDialog.image}
                onCancel={() => setImageDialog(null)}
                onApply={(png, crop) => {
                  setImageDialog(null)
                  commitBaked(imageDialog.target, png, crop)
                }}
              />
            )}
            {imageDialog?.kind === 'cutout' && (
              <CutoutDialog
                t={t}
                image={imageDialog.image}
                onCancel={() => setImageDialog(null)}
                onApply={(png) => {
                  setImageDialog(null)
                  commitBaked(imageDialog.target, png)
                }}
              />
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
                onCancel={() => {
                  setSignDlg(false)
                  setSignatureTarget(null)
                }}
                onConfirm={(sig) => {
                  setSignDlg(false)
                  if (signatureTarget) placeSignatureInField(sig, signatureTarget)
                  else setPendingSign(sig)
                }}
              />
            )}
            {staticTextDialog && (
              <div
                className="pdf-modal-mask"
                onClick={() => {
                  setStaticTextDialog(false)
                  setStaticTextEditTarget(null)
                  setTextInsertEditId(null)
                }}
              >
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">
                    {t(
                      staticTextPurpose === 'insert'
                        ? textInsertEditId
                          ? 'editInsertedText'
                          : 'insertTextTitle'
                        : staticTextEditTarget
                          ? 'formEditText'
                          : 'formAddTextTitle',
                    )}
                  </div>
                  <textarea
                    className="pdf-modal-textarea"
                    value={staticText}
                    placeholder={t('formAddTextPlaceholder')}
                    autoFocus
                    onChange={(e) => setStaticText(e.target.value)}
                  />
                  <label className="pdf-field">
                    <span>{t('formTextSize')}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={6}
                      max={72}
                      value={staticTextSize}
                      onChange={(e) =>
                        setStaticTextSize(Math.min(72, Math.max(6, Number(e.target.value) || 14)))
                      }
                    />
                  </label>
                  <div className="pdf-field-grid">
                    <div ref={staticColorFieldRef} className="pdf-field pdf-color-field">
                      <span>{t('formTextColor')}</span>
                      <button
                        type="button"
                        className="pdf-color-trigger"
                        aria-expanded={staticTextColorOpen}
                        onClick={() => setStaticTextColorOpen((open) => !open)}
                      >
                        <span
                          className="pdf-color-trigger-swatch"
                          style={{ background: staticTextColor }}
                        />
                        <span>{staticTextColor.toUpperCase()}</span>
                      </button>
                      {staticTextColorOpen && (
                        <ColorPickerPopover
                          value={staticTextColor}
                          onPick={setStaticTextColor}
                          onClose={() => setStaticTextColorOpen(false)}
                        />
                      )}
                    </div>
                    <label className="pdf-field">
                      <span>{t('formTextAlign')}</span>
                      <select
                        className="pdf-modal-input"
                        value={staticTextAlign}
                        onChange={(e) =>
                          setStaticTextAlign(e.target.value as 'left' | 'center' | 'right')
                        }
                      >
                        <option value="left">{t('formAlignLeft')}</option>
                        <option value="center">{t('formAlignCenter')}</option>
                        <option value="right">{t('formAlignRight')}</option>
                      </select>
                    </label>
                  </div>
                  <div className="pdf-modal-actions">
                    <button
                      className="pdf-modal-btn"
                      onClick={() => {
                        setStaticTextDialog(false)
                        setStaticTextEditTarget(null)
                        setTextInsertEditId(null)
                      }}
                    >
                      {t('cancel')}
                    </button>
                    <button
                      className="pdf-modal-btn primary"
                      disabled={!staticText.trim()}
                      onClick={confirmStaticFormText}
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
            {splitDlg && (
              <div className="pdf-modal-mask" onClick={() => setSplitDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('splitDlgTitle')}</div>
                  <input
                    className={`pdf-modal-input${splitInvalid ? ' invalid' : ''}`}
                    value={splitInput}
                    placeholder={t('splitDlgHint', { total: pageCount })}
                    autoFocus
                    onChange={(e) => {
                      setSplitInput(e.target.value)
                      setSplitInvalid(false)
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') confirmSplit()
                      else if (e.key === 'Escape') setSplitDlg(false)
                    }}
                  />
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setSplitDlg(false)}>
                      {t('cancel')}
                    </button>
                    <button className="pdf-modal-btn primary" onClick={confirmSplit}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {mergePagesDlg && (
              <div className="pdf-modal-mask" onClick={() => setMergePagesDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('mergePages')}</div>
                  <div className="pdf-modal-hint">{t('mergePagesHint', { total: pageCount })}</div>
                  <label className="pdf-modal-row">
                    <span>{t('mergePagesCount')}</span>
                    <input
                      className={`pdf-modal-input${mergeCountInvalid ? ' invalid' : ''}`}
                      value={mergeCount}
                      autoFocus
                      onChange={(e) => {
                        setMergeCount(e.target.value)
                        setMergeCountInvalid(false)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') confirmMergePages()
                        else if (e.key === 'Escape') setMergePagesDlg(false)
                      }}
                    />
                  </label>
                  <div className="pdf-modal-row">
                    <span>{t('mergePagesSeq')}</span>
                    <label className="pdf-modal-check">
                      <input
                        type="radio"
                        name="merge-seq"
                        checked={mergeDirection === 'horizontal'}
                        onChange={() => setMergeDirection('horizontal')}
                      />
                      {t('seqHorizontal')}
                    </label>
                    <label className="pdf-modal-check">
                      <input
                        type="radio"
                        name="merge-seq"
                        checked={mergeDirection === 'vertical'}
                        onChange={() => setMergeDirection('vertical')}
                      />
                      {t('seqVertical')}
                    </label>
                  </div>
                  <label className="pdf-modal-check">
                    <input
                      type="checkbox"
                      checked={mergeSeparator}
                      onChange={(e) => setMergeSeparator(e.target.checked)}
                    />
                    {t('mergePagesSeparator')}
                  </label>
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setMergePagesDlg(false)}>
                      {t('cancel')}
                    </button>
                    <button className="pdf-modal-btn primary" onClick={confirmMergePages}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {replaceDlg && (
              <div className="pdf-modal-mask" onClick={() => setReplaceDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('replacePages')}</div>
                  <input
                    className={`pdf-modal-input${replaceInvalid ? ' invalid' : ''}`}
                    value={replaceInput}
                    placeholder={t('replaceRangeHint', { total: pageCount })}
                    autoFocus
                    onChange={(e) => {
                      setReplaceInput(e.target.value)
                      setReplaceInvalid(false)
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') confirmReplace()
                      else if (e.key === 'Escape') setReplaceDlg(false)
                    }}
                  />
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setReplaceDlg(false)}>
                      {t('cancel')}
                    </button>
                    <button className="pdf-modal-btn primary" onClick={confirmReplace}>
                      {t('ok')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {pageSizeDlg && (
              <div className="pdf-modal-mask" onClick={() => setPageSizeDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('pageSize')}</div>
                  <div className="pdf-modal-hint">{t('pageSizeHint', { total: pageCount })}</div>
                  <div className="pdf-modal-actions merge-pages-options">
                    {PAPER_SIZES.slice(0, 3).map((p) => (
                      <button
                        key={p.label}
                        className="pdf-modal-btn"
                        onClick={() => applyPageSize(p.w, p.h)}
                      >
                        {p.label}
                      </button>
                    ))}
                  </div>
                  <div className="pdf-modal-actions merge-pages-options">
                    {PAPER_SIZES.slice(3).map((p) => (
                      <button
                        key={p.label}
                        className="pdf-modal-btn"
                        onClick={() => applyPageSize(p.w, p.h)}
                      >
                        {p.label}
                      </button>
                    ))}
                  </div>
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setPageSizeDlg(false)}>
                      {t('cancel')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {splitPagesDlg && (
              <div className="pdf-modal-mask" onClick={() => setSplitPagesDlg(false)}>
                <div className="pdf-modal" onClick={(e) => e.stopPropagation()}>
                  <div className="pdf-modal-title">{t('splitPages')}</div>
                  <div className="pdf-modal-hint">{t('splitPagesHint', { total: pageCount })}</div>
                  <div className="pdf-modal-actions merge-pages-options">
                    {([2, 4, 9] as const).map((n) => (
                      <button key={n} className="pdf-modal-btn" onClick={() => runSplitPages(n)}>
                        {t('splitNLabel', { n })}
                      </button>
                    ))}
                  </div>
                  <div className="pdf-modal-actions">
                    <button className="pdf-modal-btn" onClick={() => setSplitPagesDlg(false)}>
                      {t('cancel')}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {pageCropDlg && (
              <CropDialog
                t={t}
                title={t('cropPages')}
                image={pageCropDlg.png}
                extraFooter={
                  <label className="pdf-modal-check">
                    <input
                      type="checkbox"
                      checked={cropAllPages}
                      onChange={(e) => setCropAllPages(e.target.checked)}
                    />
                    {t('cropAllPages')}
                  </label>
                }
                onApply={(_png, crop) => confirmPageCrop(crop)}
                onCancel={() => setPageCropDlg(null)}
              />
            )}
          </div>

          <footer className="status-bar">
            <div className="status-left">
              <span className="status-item">
                {t('appPageOf', { current: currentPage, total: pageCount })}
              </span>
            </div>
            <div className="status-right">
              <button
                className="zoom-btn"
                data-tip={t('zoomOut')}
                aria-label={t('zoomOut')}
                onClick={zoomOut}
              >
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
              <button
                className="zoom-btn"
                data-tip={t('zoomIn')}
                aria-label={t('zoomIn')}
                onClick={zoomIn}
              >
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
