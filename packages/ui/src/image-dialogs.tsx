/**
 * Picture pixel-edit dialogs shared by the office apps: remove background (tolerance cutout)
 * and crop. Ported from apps/pdf ImageDialogs (itself a port of docs' PictureDialogs); apps
 * pass their own strings through `labels` and import `image-dialogs.css` for the chrome.
 *
 * Both take the source as a data URL and hand back a base64 PNG (no data: prefix).
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import type {
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
  ReactElement,
  ReactNode,
} from 'react'
import { removeBackground, sampleBackgroundColors, type PixelImage, type RGB } from './cutout'
import { useModalKeys } from './modal-keys'
import { CROP_EDGE_LABELS } from './strings-crop-edges'
import type { Lang } from '@genoffice/i18n'

export interface ImageDialogLabels {
  cancel: string
  apply: string
  applying: string
  loading: string
  loadFailed: string
  processFailed: string
  cutoutTitle: string
  tolerance: string
  cutoutHint: (removedPct: number) => string
  cropTitle: string
  cropHint: string
}

/** Kept region of a crop, as 0..1 fractions of the image (l<r, t<b) */
export interface CropFractions {
  l: number
  t: number
  r: number
  b: number
}

/** Remove-background tolerance (0..100) the slider starts from */
export const DEFAULT_CUTOUT_TOLERANCE = 30

/** Longest side of the preview canvas (px) */
const PREVIEW_MAX = 520
const CROP_HANDLE_GUTTER = 6

const fitPreview = (
  naturalWidth: number,
  naturalHeight: number,
  maxWidth: number,
  maxHeight: number,
): { w: number; h: number } => {
  const scale = Math.min(1, maxWidth / naturalWidth, maxHeight / naturalHeight)
  return {
    w: Math.max(1, Math.round(naturalWidth * scale)),
    h: Math.max(1, Math.round(naturalHeight * scale)),
  }
}

/** Checkerboard backdrop visualising transparent areas (document-neutral, not themed) */
const CHECKERBOARD: CSSProperties = {
  background: 'repeating-conic-gradient(#d5d5d5 0% 25%, #ffffff 0% 50%) 0 0 / 16px 16px',
}

const toBase64 = (dataUrl: string): string => dataUrl.split(',')[1] ?? ''

/** Kept region of a decoded image as a base64 PNG (throws on canvas failure) */
export function cropImagePng(img: HTMLImageElement, crop: CropFractions): string {
  const w = img.naturalWidth
  const h = img.naturalHeight
  const sx = Math.round(crop.l * w)
  const sy = Math.round(crop.t * h)
  const sw = Math.max(1, Math.round((crop.r - crop.l) * w))
  const sh = Math.max(1, Math.round((crop.b - crop.t) * h))
  const c = document.createElement('canvas')
  c.width = sw
  c.height = sh
  c.getContext('2d')!.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh)
  return toBase64(c.toDataURL('image/png'))
}

type ErrorKey = 'loadFailed' | 'processFailed'

function Stage({
  labels,
  error,
  loaded,
  style,
  children,
}: {
  labels: ImageDialogLabels
  error: ErrorKey | null
  loaded: boolean
  style?: CSSProperties
  children: ReactNode
}) {
  return (
    <div className="gs-imgdlg-stage" style={{ ...CHECKERBOARD, ...style }}>
      {error ? <span className="gs-imgdlg-error">{labels[error]}</span> : children}
      {!loaded && !error && <span className="gs-imgdlg-dim">{labels.loading}</span>}
    </div>
  )
}

/* ================= Remove background ================= */

export function CutoutDialog({
  labels,
  image,
  onApply,
  onCancel,
}: {
  labels: ImageDialogLabels
  /** source picture as a data URL */
  image: string
  /** background-removed PNG (base64, no data: prefix) */
  onApply: (png: string) => void
  onCancel: () => void
}): ReactElement {
  const [tolerance, setTolerance] = useState(DEFAULT_CUTOUT_TOLERANCE)
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState<ErrorKey | null>(null)
  const [removedPct, setRemovedPct] = useState(0)
  const [applying, setApplying] = useState(false)

  const canvasRef = useRef<HTMLCanvasElement>(null)
  /** source-resolution pixels (used on apply) */
  const fullRef = useRef<PixelImage | null>(null)
  /** downsampled pixels for the slider's live preview */
  const previewRef = useRef<PixelImage | null>(null)
  /** background representative colours, sampled once at source resolution */
  const bgColorsRef = useRef<RGB[]>([])
  const rafRef = useRef<number | null>(null)

  const renderPreview = useCallback((tol: number) => {
    const pv = previewRef.current
    const canvas = canvasRef.current
    if (!pv || !canvas) return
    const result = removeBackground(pv, tol, bgColorsRef.current)
    canvas.getContext('2d')!.putImageData(new ImageData(result.data, pv.width, pv.height), 0, 0)
    setRemovedPct(Math.round((result.removedCount / (pv.width * pv.height)) * 100))
  }, [])

  useEffect(() => {
    let cancelled = false
    const img = new Image()
    img.onload = () => {
      if (cancelled) return
      const w = img.naturalWidth
      const h = img.naturalHeight
      if (!w || !h) {
        setError('loadFailed')
        return
      }
      const grab = (dw: number, dh: number): PixelImage => {
        const c = document.createElement('canvas')
        c.width = dw
        c.height = dh
        const ctx = c.getContext('2d')!
        ctx.drawImage(img, 0, 0, dw, dh)
        const d = ctx.getImageData(0, 0, dw, dh)
        return { data: d.data, width: dw, height: dh }
      }
      const full = grab(w, h)
      const scale = Math.min(1, PREVIEW_MAX / Math.max(w, h))
      const preview =
        scale < 1
          ? grab(Math.max(1, Math.round(w * scale)), Math.max(1, Math.round(h * scale)))
          : full
      fullRef.current = full
      previewRef.current = preview
      bgColorsRef.current = sampleBackgroundColors(full)
      const canvas = canvasRef.current
      if (canvas) {
        canvas.width = preview.width
        canvas.height = preview.height
      }
      setLoaded(true)
      renderPreview(DEFAULT_CUTOUT_TOLERANCE)
    }
    img.onerror = () => {
      if (!cancelled) setError('loadFailed')
    }
    img.src = image
    return () => {
      cancelled = true
    }
  }, [image, renderPreview])

  const onTolerance = (v: number) => {
    setTolerance(v)
    if (rafRef.current != null) cancelAnimationFrame(rafRef.current)
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = null
      renderPreview(v)
    })
  }
  useEffect(
    () => () => {
      if (rafRef.current != null) cancelAnimationFrame(rafRef.current)
    },
    [],
  )

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        onCancel()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onCancel])

  const apply = () => {
    const full = fullRef.current
    if (!full || applying) return
    setApplying(true)
    // let the "processing…" state paint before the heavy recompute
    window.setTimeout(() => {
      try {
        const result = removeBackground(full, tolerance, bgColorsRef.current)
        const c = document.createElement('canvas')
        c.width = full.width
        c.height = full.height
        c.getContext('2d')!.putImageData(new ImageData(result.data, full.width, full.height), 0, 0)
        onApply(toBase64(c.toDataURL('image/png')))
      } catch {
        setApplying(false)
        setError('processFailed')
      }
    }, 30)
  }

  return (
    <div className="gs-imgdlg-mask" onClick={onCancel}>
      <div
        className="gs-imgdlg"
        role="dialog"
        aria-label={labels.cutoutTitle}
        style={{ maxWidth: PREVIEW_MAX + 48 }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="gs-imgdlg-title">{labels.cutoutTitle}</div>
        <Stage labels={labels} error={error} loaded={loaded} style={{ maxHeight: PREVIEW_MAX }}>
          <canvas
            ref={canvasRef}
            style={{ maxWidth: '100%', maxHeight: PREVIEW_MAX, display: loaded ? 'block' : 'none' }}
          />
        </Stage>
        <label className="gs-imgdlg-slider">
          <span>{labels.tolerance}</span>
          <input
            type="range"
            min={0}
            max={100}
            step={1}
            value={tolerance}
            disabled={!loaded || applying}
            onChange={(e) => onTolerance(Number(e.target.value))}
          />
          <span className="gs-imgdlg-slider-value">{tolerance}</span>
        </label>
        <div className="gs-imgdlg-hint">{labels.cutoutHint(removedPct)}</div>
        <div className="gs-imgdlg-actions">
          <button type="button" className="gs-imgdlg-btn" onClick={onCancel} disabled={applying}>
            {labels.cancel}
          </button>
          <button
            type="button"
            className="gs-imgdlg-btn primary"
            onClick={apply}
            disabled={!loaded || !!error || applying}
          >
            {applying ? labels.applying : labels.apply}
          </button>
        </div>
      </div>
    </div>
  )
}

/* ================= Crop ================= */

type CropHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'move'

export type CropEdge = Exclude<CropHandle, 'move'>

/** minimum crop-box side (preview px), prevents dragging to 0 */
const MIN_CROP_PX = 16
const HANDLES: CropEdge[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
export const CROP_EDGES = HANDLES
/** keyboard edge nudge as a fraction of the image; Shift takes the coarse step */
export const CROP_EDGE_STEP = 0.01
export const CROP_EDGE_STEP_COARSE = 0.1

/**
 * Move one crop edge by (dx, dy) fractions of the image. Shares the drag
 * clamp so a keyboard nudge and a pointer drag cannot disagree: each side is
 * limited to the image bounds and to MIN_CROP_PX away from its opposite.
 */
export function nudgeCropEdge(
  crop: CropFractions,
  handle: CropEdge,
  dx: number,
  dy: number,
  minW: number,
  minH: number,
): CropFractions {
  let { l, t, r, b } = crop
  if (handle.includes('w')) l = Math.min(Math.max(0, l + dx), r - minW)
  if (handle.includes('e')) r = Math.max(Math.min(1, r + dx), l + minW)
  if (handle.includes('n')) t = Math.min(Math.max(0, t + dy), b - minH)
  if (handle.includes('s')) b = Math.max(Math.min(1, b + dy), t + minH)
  return { l, t, r, b }
}

/** Fraction a handle reports as its slider value: the vertical side for a
 * top/bottom edge (including corners), the horizontal side otherwise. */
export function cropEdgeValue(crop: CropFractions, handle: CropEdge): number {
  if (handle.includes('n')) return crop.t
  if (handle.includes('s')) return crop.b
  if (handle.includes('w')) return crop.l
  return crop.r
}

/** Arrow key to edge nudge, or null when the key does not move this edge. */
export function cropEdgeArrowDelta(
  key: string,
  handle: CropEdge,
  coarse: boolean,
): { dx: number; dy: number } | null {
  const step = coarse ? CROP_EDGE_STEP_COARSE : CROP_EDGE_STEP
  const dx = key === 'ArrowLeft' ? -step : key === 'ArrowRight' ? step : 0
  const dy = key === 'ArrowUp' ? -step : key === 'ArrowDown' ? step : 0
  if (dx === 0 && dy === 0) return null
  const movesX = handle.includes('w') || handle.includes('e')
  const movesY = handle.includes('n') || handle.includes('s')
  if (dx !== 0 && !movesX) return null
  if (dy !== 0 && !movesY) return null
  return { dx: movesX ? dx : 0, dy: movesY ? dy : 0 }
}

const HANDLE_CURSOR: Record<string, string> = {
  nw: 'nwse-resize',
  se: 'nwse-resize',
  ne: 'nesw-resize',
  sw: 'nesw-resize',
  n: 'ns-resize',
  s: 'ns-resize',
  e: 'ew-resize',
  w: 'ew-resize',
}

function handleStyle(pos: CropHandle): CSSProperties {
  const base: CSSProperties = { cursor: HANDLE_CURSOR[pos] }
  if (pos.includes('n')) base.top = -5
  if (pos.includes('s')) base.bottom = -5
  if (pos.includes('w')) base.left = -5
  if (pos.includes('e')) base.right = -5
  if (pos === 'n' || pos === 's') {
    base.left = '50%'
    base.marginLeft = -5
  }
  if (pos === 'e' || pos === 'w') {
    base.top = '50%'
    base.marginTop = -5
  }
  return base
}

export function CropDialog({
  labels,
  image,
  onApply,
  onCancel,
  extraFooter,
  lang = 'en',
}: {
  labels: ImageDialogLabels
  /** source picture as a data URL */
  image: string
  /** cropped PNG (base64, no data: prefix) plus the kept-region fractions */
  onApply: (png: string, crop: CropFractions) => void
  onCancel: () => void
  /** extra controls between the hint and the action buttons */
  extraFooter?: ReactNode
  /** language of the edge handle names */
  lang?: Lang
}): ReactElement {
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState<ErrorKey | null>(null)
  const [crop, setCrop] = useState<CropFractions>({ l: 0, t: 0, r: 1, b: 1 })
  /** preview display size (CSS px) */
  const [view, setView] = useState<{ w: number; h: number } | null>(null)
  const imgRef = useRef<HTMLImageElement | null>(null)
  const dragRef = useRef<{
    handle: CropHandle
    startX: number
    startY: number
    start: CropFractions
  } | null>(null)

  const updateView = useCallback(() => {
    const img = imgRef.current
    if (!img) return
    const maxWidth = Math.max(80, Math.min(PREVIEW_MAX, window.innerWidth - 92))
    const maxHeight = Math.max(80, Math.min(PREVIEW_MAX, window.innerHeight - 220))
    setView(fitPreview(img.naturalWidth, img.naturalHeight, maxWidth, maxHeight))
  }, [])

  useEffect(() => {
    window.addEventListener('resize', updateView)
    return () => window.removeEventListener('resize', updateView)
  }, [updateView])

  useEffect(() => {
    let cancelled = false
    const img = new Image()
    img.onload = () => {
      if (cancelled) return
      if (!img.naturalWidth || !img.naturalHeight) {
        setError('loadFailed')
        return
      }
      imgRef.current = img
      updateView()
      setLoaded(true)
    }
    img.onerror = () => {
      if (!cancelled) setError('loadFailed')
    }
    img.src = image
    return () => {
      cancelled = true
    }
  }, [image, updateView])

  const apply = useCallback(() => {
    const img = imgRef.current
    if (!img) return
    try {
      onApply(cropImagePng(img, crop), crop)
    } catch {
      setError('processFailed')
    }
  }, [crop, onApply])

  // Esc / Tab / initial focus come from useModalKeys on the backdrop.
  const modalKeys = useModalKeys(onCancel, { restoreFocus: true })
  // useModalKeys lands on the first control, which is Cancel, so Enter would
  // cancel while the hint promises "Enter to apply". Apply is disabled until
  // the picture is in; once it is, hand that default focus to Apply unless the
  // user already moved on to a handle.
  const applyRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    if (!loaded || error) return
    const backdrop = modalKeys.ref.current
    const active = document.activeElement
    const first = backdrop?.querySelector('input, textarea, select, button')
    if (!backdrop?.contains(active) || active === backdrop || active === first) {
      applyRef.current?.focus()
    }
  }, [loaded, error, modalKeys.ref])

  const minSide = useCallback(
    () => ({
      w: view ? MIN_CROP_PX / view.w : 0,
      h: view ? MIN_CROP_PX / view.h : 0,
    }),
    [view],
  )

  const onHandleKeyDown = (e: ReactKeyboardEvent, handle: CropEdge) => {
    const delta = cropEdgeArrowDelta(e.key, handle, e.shiftKey)
    if (!delta) return
    const min = minSide()
    const next = nudgeCropEdge(crop, handle, delta.dx, delta.dy, min.w, min.h)
    if (next.l === crop.l && next.t === crop.t && next.r === crop.r && next.b === crop.b) return
    e.preventDefault()
    e.stopPropagation()
    setCrop(next)
  }

  const startDrag = (handle: CropHandle) => (e: ReactMouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    dragRef.current = { handle, startX: e.clientX, startY: e.clientY, start: crop }
  }

  useEffect(() => {
    if (!view) return
    const minW = MIN_CROP_PX / view.w
    const minH = MIN_CROP_PX / view.h
    const onMove = (e: MouseEvent) => {
      const drag = dragRef.current
      if (!drag) return
      const dx = (e.clientX - drag.startX) / view.w
      const dy = (e.clientY - drag.startY) / view.h
      const s = drag.start
      let { l, t: top, r, b } = s
      if (drag.handle === 'move') {
        const w = s.r - s.l
        const h = s.b - s.t
        l = Math.min(Math.max(0, s.l + dx), 1 - w)
        top = Math.min(Math.max(0, s.t + dy), 1 - h)
        r = l + w
        b = top + h
      } else {
        if (drag.handle.includes('w')) l = Math.min(Math.max(0, s.l + dx), s.r - minW)
        if (drag.handle.includes('e')) r = Math.max(Math.min(1, s.r + dx), s.l + minW)
        if (drag.handle.includes('n')) top = Math.min(Math.max(0, s.t + dy), s.b - minH)
        if (drag.handle.includes('s')) b = Math.max(Math.min(1, s.b + dy), s.t + minH)
      }
      setCrop({ l, t: top, r, b })
    }
    const onUp = () => {
      dragRef.current = null
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [view])

  const px = (v: number, total: number) => Math.round(v * total)

  return (
    <div
      className="gs-imgdlg-mask"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onClick={onCancel}
    >
      <div
        className="gs-imgdlg gs-imgdlg-crop"
        role="dialog"
        aria-modal="true"
        aria-label={labels.cropTitle}
        style={{ width: PREVIEW_MAX + 48 + CROP_HANDLE_GUTTER * 2 }}
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key !== 'Enter' || !loaded || error) return
          // a focused edge or button owns Enter: a global Enter must not apply
          // a crop the user did not ask for
          if ((e.target as HTMLElement | null)?.closest?.('button, [role="slider"]')) return
          e.preventDefault()
          apply()
        }}
      >
        <div className="gs-imgdlg-title">{labels.cropTitle}</div>
        <Stage
          labels={labels}
          error={error}
          loaded={loaded}
          style={{ padding: CROP_HANDLE_GUTTER }}
        >
          {loaded && view && (
            <div className="gs-imgdlg-crop-view" style={{ width: view.w, height: view.h }}>
              <img src={image} width={view.w} height={view.h} draggable={false} alt="" />
              {/* the four dimmed regions outside the crop box */}
              <div
                className="gs-imgdlg-shade"
                style={{ left: 0, top: 0, width: '100%', height: px(crop.t, view.h) }}
              />
              <div
                className="gs-imgdlg-shade"
                style={{ left: 0, bottom: 0, width: '100%', height: view.h - px(crop.b, view.h) }}
              />
              <div
                className="gs-imgdlg-shade"
                style={{
                  left: 0,
                  top: px(crop.t, view.h),
                  width: px(crop.l, view.w),
                  height: px(crop.b, view.h) - px(crop.t, view.h),
                }}
              />
              <div
                className="gs-imgdlg-shade"
                style={{
                  right: 0,
                  top: px(crop.t, view.h),
                  width: view.w - px(crop.r, view.w),
                  height: px(crop.b, view.h) - px(crop.t, view.h),
                }}
              />
              <div
                className="gs-imgdlg-crop-box"
                style={{
                  left: `${crop.l * 100}%`,
                  top: `${crop.t * 100}%`,
                  width: `${(crop.r - crop.l) * 100}%`,
                  height: `${(crop.b - crop.t) * 100}%`,
                }}
                onMouseDown={startDrag('move')}
              >
                {HANDLES.map((pos) => (
                  <div
                    key={pos}
                    className="gs-imgdlg-handle"
                    style={handleStyle(pos)}
                    role="slider"
                    tabIndex={0}
                    aria-label={CROP_EDGE_LABELS[lang][pos]}
                    aria-valuemin={0}
                    aria-valuemax={100}
                    aria-valuenow={Math.round(cropEdgeValue(crop, pos) * 100)}
                    aria-orientation={
                      pos.includes('n') || pos.includes('s') ? 'vertical' : 'horizontal'
                    }
                    data-crop-edge={pos}
                    onMouseDown={startDrag(pos)}
                    onKeyDown={(e) => onHandleKeyDown(e, pos)}
                  />
                ))}
              </div>
            </div>
          )}
        </Stage>
        <div className="gs-imgdlg-hint">{labels.cropHint}</div>
        {extraFooter}
        <div className="gs-imgdlg-actions">
          <button type="button" className="gs-imgdlg-btn" onClick={onCancel}>
            {labels.cancel}
          </button>
          <button
            type="button"
            className="gs-imgdlg-btn primary"
            ref={applyRef}
            onClick={apply}
            disabled={!loaded || !!error}
          >
            {labels.apply}
          </button>
        </div>
      </div>
    </div>
  )
}
