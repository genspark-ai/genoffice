/**
 * Picture Format dialogs: remove background (tolerance cutout) + crop.
 *
 * Remove background: same interaction as slides' CutoutDialog — a tolerance slider with live
 * preview; the preview computes on a ≤520px downsampled copy, applying recomputes at the
 * original resolution and returns a transparent PNG dataUrl.
 *
 * Crop: drag 8 handles on the preview (or drag inside the box to move) to pick the kept
 * region; applying bakes the selection into a new dataUrl at the original resolution
 * (png keeps transparency, jpeg stays jpeg). Reuses the .modal-backdrop/.modal styles.
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'
import { removeBackground, sampleBackgroundColors, type PixelImage, type RGB } from '../cutout'
import { useI18n, type StringKey } from '../i18n/locale'

/** Longest side of the preview canvas (px) */
const PREVIEW_MAX = 520
const DEFAULT_TOLERANCE = 30
const CROP_HANDLE_GUTTER = 6

export function fitCropPreview(
  naturalWidth: number,
  naturalHeight: number,
  maxWidth = PREVIEW_MAX,
  maxHeight = PREVIEW_MAX,
): { w: number; h: number } {
  const scale = Math.min(1, maxWidth / naturalWidth, maxHeight / naturalHeight)
  return {
    w: Math.max(1, Math.round(naturalWidth * scale)),
    h: Math.max(1, Math.round(naturalHeight * scale)),
  }
}

/** Checkerboard backdrop: visualizes transparent areas */
const CHECKERBOARD: React.CSSProperties = {
  background: 'repeating-conic-gradient(#d5d5d5 0% 25%, #ffffff 0% 50%) 0 0 / 16px 16px',
}

/* ================= Remove background ================= */

interface CutoutProps {
  dataUrl: string
  /** Apply: the background-removed PNG dataUrl (with alpha channel) */
  onApply: (pngDataUrl: string) => void
  onCancel: () => void
}

export function CutoutDialog({ dataUrl, onApply, onCancel }: CutoutProps) {
  const { t } = useI18n()
  const [tolerance, setTolerance] = useState(DEFAULT_TOLERANCE)
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState<StringKey | null>(null)
  const [removedPct, setRemovedPct] = useState(0)
  const [applying, setApplying] = useState(false)

  const canvasRef = useRef<HTMLCanvasElement>(null)
  /** Original-resolution pixels (used on apply) */
  const fullRef = useRef<PixelImage | null>(null)
  /** Downsampled pixels (for the slider's live preview) */
  const previewRef = useRef<PixelImage | null>(null)
  /** Background representative colors (from original-resolution border sampling, shared by preview/apply for consistency) */
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
        setError('ribbonImageLoadFail')
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
      renderPreview(DEFAULT_TOLERANCE)
    }
    img.onerror = () => {
      if (!cancelled) setError('ribbonImageLoadFail')
    }
    img.src = dataUrl
    return () => {
      cancelled = true
    }
  }, [dataUrl, renderPreview])

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
    // let the "processing…" state render before the heavy recompute
    window.setTimeout(() => {
      try {
        const result = removeBackground(full, tolerance, bgColorsRef.current)
        const c = document.createElement('canvas')
        c.width = full.width
        c.height = full.height
        c.getContext('2d')!.putImageData(new ImageData(result.data, full.width, full.height), 0, 0)
        onApply(c.toDataURL('image/png'))
      } catch {
        setApplying(false)
        setError('ribbonImageProcessFail')
      }
    }, 30)
  }

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div
        className="modal"
        style={{ maxWidth: PREVIEW_MAX + 48 }}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('ribbonRemoveBg')}</h2>
        <div
          style={{
            ...CHECKERBOARD,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: 160,
            maxHeight: PREVIEW_MAX,
            borderRadius: 4,
            overflow: 'hidden',
          }}
        >
          {error ? (
            <span
              style={{
                color: 'var(--danger)',
                background: 'var(--surface)',
                padding: '4px 10px',
                borderRadius: 4,
              }}
            >
              {t(error)}
            </span>
          ) : (
            <canvas
              ref={canvasRef}
              style={{
                maxWidth: '100%',
                maxHeight: PREVIEW_MAX,
                display: loaded ? 'block' : 'none',
              }}
            />
          )}
          {!loaded && !error && (
            <span style={{ color: 'var(--text-muted)' }}>{t('ribbonLoading')}</span>
          )}
        </div>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 12 }}>
          <span style={{ whiteSpace: 'nowrap' }}>{t('ribbonCutoutTolerance')}</span>
          <input
            type="range"
            min={0}
            max={100}
            step={1}
            value={tolerance}
            disabled={!loaded || applying}
            onChange={(e) => onTolerance(Number(e.target.value))}
            style={{ flex: 1 }}
          />
          <span style={{ width: 32, textAlign: 'right' }}>{tolerance}</span>
        </label>
        <div style={{ fontSize: 12, color: 'var(--text-muted)', marginTop: 4 }}>
          {t('ribbonCutoutHint', { pct: removedPct })}
        </div>
        <div className="modal-actions">
          <button onClick={onCancel} disabled={applying}>
            {t('ribbonCancel')}
          </button>
          <button className="primary" onClick={apply} disabled={!loaded || !!error || applying}>
            {applying ? t('ribbonProcessing') : t('ribbonApply')}
          </button>
        </div>
      </div>
    </div>
  )
}

/* ================= Crop ================= */

/** Crop region (0..1 ratios, relative to the whole image) */
interface CropRect {
  l: number
  t: number
  r: number
  b: number
}

type CropHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'move'

/** Minimum crop-box side (preview px), prevents dragging to 0 */
const MIN_CROP_PX = 16

interface CropProps {
  dataUrl: string
  /** Apply: the cropped dataUrl (png/jpeg follows the original format) */
  onApply: (croppedDataUrl: string) => void
  onCancel: () => void
}

export function CropDialog({ dataUrl, onApply, onCancel }: CropProps) {
  const { t } = useI18n()
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState<StringKey | null>(null)
  const [crop, setCrop] = useState<CropRect>({ l: 0, t: 0, r: 1, b: 1 })
  /** Preview display size (CSS px) */
  const [view, setView] = useState<{ w: number; h: number } | null>(null)
  const imgRef = useRef<HTMLImageElement | null>(null)
  const boxRef = useRef<HTMLDivElement>(null)
  const dragRef = useRef<{
    handle: CropHandle
    startX: number
    startY: number
    start: CropRect
  } | null>(null)

  const updateView = useCallback(() => {
    const img = imgRef.current
    if (!img) return
    const maxWidth = Math.max(80, Math.min(PREVIEW_MAX, window.innerWidth - 92))
    const maxHeight = Math.max(80, Math.min(PREVIEW_MAX, window.innerHeight - 220))
    setView(fitCropPreview(img.naturalWidth, img.naturalHeight, maxWidth, maxHeight))
  }, [])

  useEffect(() => {
    document.body.classList.add('docs-crop-active')
    return () => document.body.classList.remove('docs-crop-active')
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
        setError('ribbonImageLoadFail')
        return
      }
      imgRef.current = img
      updateView()
      setLoaded(true)
    }
    img.onerror = () => {
      if (!cancelled) setError('ribbonImageLoadFail')
    }
    img.src = dataUrl
    return () => {
      cancelled = true
    }
  }, [dataUrl, updateView])

  const apply = useCallback(() => {
    const img = imgRef.current
    if (!img) return
    const w = img.naturalWidth
    const h = img.naturalHeight
    const sx = Math.round(crop.l * w)
    const sy = Math.round(crop.t * h)
    const sw = Math.max(1, Math.round((crop.r - crop.l) * w))
    const sh = Math.max(1, Math.round((crop.b - crop.t) * h))
    try {
      const c = document.createElement('canvas')
      c.width = sw
      c.height = sh
      c.getContext('2d')!.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh)
      // jpeg has no alpha channel; keep jpeg to avoid size bloat. Others (png/gif) all output png to preserve transparency
      const isJpeg = /^data:image\/jpeg/.test(dataUrl)
      onApply(isJpeg ? c.toDataURL('image/jpeg', 0.92) : c.toDataURL('image/png'))
    } catch {
      setError('ribbonImageProcessFail')
    }
  }, [crop, dataUrl, onApply])

  // Esc cancels / Enter applies
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        onCancel()
      } else if (e.key === 'Enter' && loaded && !error) {
        e.preventDefault()
        apply()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onCancel, apply, loaded, error])

  const startDrag = (handle: CropHandle) => (e: React.MouseEvent) => {
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
      let { l, t, r, b } = s
      if (drag.handle === 'move') {
        const w = s.r - s.l
        const h = s.b - s.t
        l = Math.min(Math.max(0, s.l + dx), 1 - w)
        t = Math.min(Math.max(0, s.t + dy), 1 - h)
        r = l + w
        b = t + h
      } else {
        if (drag.handle.includes('w')) l = Math.min(Math.max(0, s.l + dx), s.r - minW)
        if (drag.handle.includes('e')) r = Math.max(Math.min(1, s.r + dx), s.l + minW)
        if (drag.handle.includes('n')) t = Math.min(Math.max(0, s.t + dy), s.b - minH)
        if (drag.handle.includes('s')) b = Math.max(Math.min(1, s.b + dy), s.t + minH)
      }
      setCrop({ l, t, r, b })
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

  /** Shared handle style */
  const handleStyle = (pos: CropHandle): React.CSSProperties => {
    const base: React.CSSProperties = {
      position: 'absolute',
      width: 10,
      height: 10,
      background: '#fff',
      border: '1.5px solid #2b5797',
      borderRadius: 2,
      boxSizing: 'border-box',
      zIndex: 2,
    }
    const cursors: Record<string, string> = {
      nw: 'nwse-resize',
      se: 'nwse-resize',
      ne: 'nesw-resize',
      sw: 'nesw-resize',
      n: 'ns-resize',
      s: 'ns-resize',
      e: 'ew-resize',
      w: 'ew-resize',
    }
    base.cursor = cursors[pos]
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

  const HANDLES: CropHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div
        className="modal crop-modal"
        style={{
          width: PREVIEW_MAX + 48 + CROP_HANDLE_GUTTER * 2,
          maxWidth: 'calc(100vw - 32px)',
          maxHeight: 'calc(100vh - 32px)',
          boxSizing: 'border-box',
          overflowY: 'auto',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('ribbonCrop')}</h2>
        <div
          style={{
            ...CHECKERBOARD,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: 160,
            borderRadius: 4,
            padding: CROP_HANDLE_GUTTER,
            boxSizing: 'border-box',
            overflow: 'hidden',
          }}
        >
          {error ? (
            <span
              style={{
                color: 'var(--danger)',
                background: 'var(--surface)',
                padding: '4px 10px',
                borderRadius: 4,
              }}
            >
              {t(error)}
            </span>
          ) : loaded && view ? (
            <div
              ref={boxRef}
              style={{ position: 'relative', width: view.w, height: view.h, userSelect: 'none' }}
            >
              <img
                src={dataUrl}
                width={view.w}
                height={view.h}
                draggable={false}
                style={{ display: 'block' }}
              />
              {/* the four masks outside the crop box */}
              {(() => {
                const px = (v: number, total: number) => Math.round(v * total)
                const dim: React.CSSProperties = {
                  position: 'absolute',
                  background: 'rgba(0,0,0,0.45)',
                }
                return (
                  <>
                    <div
                      style={{ ...dim, left: 0, top: 0, width: '100%', height: px(crop.t, view.h) }}
                    />
                    <div
                      style={{
                        ...dim,
                        left: 0,
                        bottom: 0,
                        width: '100%',
                        height: view.h - px(crop.b, view.h),
                      }}
                    />
                    <div
                      style={{
                        ...dim,
                        left: 0,
                        top: px(crop.t, view.h),
                        width: px(crop.l, view.w),
                        height: px(crop.b, view.h) - px(crop.t, view.h),
                      }}
                    />
                    <div
                      style={{
                        ...dim,
                        right: 0,
                        top: px(crop.t, view.h),
                        width: view.w - px(crop.r, view.w),
                        height: px(crop.b, view.h) - px(crop.t, view.h),
                      }}
                    />
                  </>
                )
              })()}
              {/* crop box + handles */}
              <div
                style={{
                  position: 'absolute',
                  left: `${crop.l * 100}%`,
                  top: `${crop.t * 100}%`,
                  width: `${(crop.r - crop.l) * 100}%`,
                  height: `${(crop.b - crop.t) * 100}%`,
                  border: '1.5px solid #2b5797',
                  boxSizing: 'border-box',
                  cursor: 'move',
                }}
                onMouseDown={startDrag('move')}
              >
                {HANDLES.map((pos) => (
                  <div key={pos} style={handleStyle(pos)} onMouseDown={startDrag(pos)} />
                ))}
              </div>
            </div>
          ) : (
            <span style={{ color: 'var(--text-muted)' }}>{t('ribbonLoading')}</span>
          )}
        </div>
        <div style={{ fontSize: 12, color: 'var(--text-muted)', marginTop: 8 }}>
          {t('ribbonCropHint')}
        </div>
        <div className="modal-actions">
          <button onClick={onCancel}>{t('ribbonCancel')}</button>
          <button className="primary" onClick={apply} disabled={!loaded || !!error}>
            {t('ribbonApply')}
          </button>
        </div>
      </div>
    </div>
  )
}
