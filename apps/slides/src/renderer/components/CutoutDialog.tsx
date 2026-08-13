/**
 * Cutout (background removal) dialog — a simplified take on PowerPoint "Remove Background".
 *
 * Entry: select a picture → format pane / "Picture Format" tab / context menu "Remove background".
 * Interaction: tolerance slider (0-100) with live preview, checkerboard background shows
 * transparent areas; on apply, calls back with the background-removed PNG dataUrl (with alpha)
 * and App replaces the original picture.
 *
 * Performance: slider preview computes on a ≤520px downsampled copy to stay smooth; clicking
 * "Apply" recomputes at the original resolution, and representative background colors are always
 * sampled from the original-resolution border.
 * Reuses SettingsModal's .modal-backdrop/.modal styles.
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'
import { removeBackground, sampleBackgroundColors, type PixelImage, type RGB } from '../cutout'
import { useI18n, type StringKey } from '../i18n/locale'

/** Longest side of the preview canvas (px) */
const PREVIEW_MAX = 520
const DEFAULT_TOLERANCE = 30

/** Checkerboard background: visualizes transparent areas */
const CHECKERBOARD: React.CSSProperties = {
  background: 'repeating-conic-gradient(#d5d5d5 0% 25%, #ffffff 0% 50%) 0 0 / 16px 16px',
}

interface Props {
  /** Original image dataUrl (full source, without crop) */
  dataUrl: string
  /** Apply: the background-removed PNG dataUrl (with alpha channel) */
  onApply: (pngDataUrl: string) => void
  onCancel: () => void
}

export function CutoutDialog({ dataUrl, onApply, onCancel }: Props) {
  const { t } = useI18n()
  const [tolerance, setTolerance] = useState(DEFAULT_TOLERANCE)
  const [loaded, setLoaded] = useState(false)
  const [error, setError] = useState<StringKey | null>(null)
  const [removedPct, setRemovedPct] = useState(0)
  const [applying, setApplying] = useState(false)

  const canvasRef = useRef<HTMLCanvasElement>(null)
  /** Original-resolution pixels (used on apply) */
  const fullRef = useRef<PixelImage | null>(null)
  /** Downsampled pixels (for live slider preview) */
  const previewRef = useRef<PixelImage | null>(null)
  /** Representative background colors (sampled from the original-resolution border; shared by preview/apply for consistency) */
  const bgColorsRef = useRef<RGB[]>([])
  const rafRef = useRef<number | null>(null)

  /** Recompute the preview at the current tolerance and draw it */
  const renderPreview = useCallback((tol: number) => {
    const pv = previewRef.current
    const canvas = canvasRef.current
    if (!pv || !canvas) return
    const result = removeBackground(pv, tol, bgColorsRef.current)
    canvas.getContext('2d')!.putImageData(new ImageData(result.data, pv.width, pv.height), 0, 0)
    setRemovedPct(Math.round((result.removedCount / (pv.width * pv.height)) * 100))
  }, [])

  // Load the original → two pixel sets (original resolution + preview resolution), and draw the first preview frame
  useEffect(() => {
    let cancelled = false
    const img = new Image()
    img.onload = () => {
      if (cancelled) return
      const w = img.naturalWidth
      const h = img.naturalHeight
      if (!w || !h) {
        setError('paneCutoutDecodeFail')
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
      if (!cancelled) setError('paneCutoutLoadFail')
    }
    img.src = dataUrl
    return () => {
      cancelled = true
    }
  }, [dataUrl, renderPreview])

  // Slider change → recompute the preview coalesced by rAF
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

  // Esc cancels
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

  /** Apply: recompute at original resolution → PNG dataUrl with alpha */
  const apply = () => {
    const full = fullRef.current
    if (!full || applying) return
    setApplying(true)
    // Let "Processing…" render first, then do the recomputation
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
        setError('paneCutoutApplyFail')
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
        <h2>{t('paneCutoutTitle')}</h2>
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
              style={{ color: '#c33', background: '#fff', padding: '4px 10px', borderRadius: 4 }}
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
          {!loaded && !error && <span style={{ color: '#888' }}>{t('paneCutoutLoading')}</span>}
        </div>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 12 }}>
          <span style={{ whiteSpace: 'nowrap' }}>{t('paneCutoutTolerance')}</span>
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
        <div style={{ fontSize: 12, color: '#888', marginTop: 4 }}>
          {t('paneCutoutHint', { pct: removedPct })}
        </div>
        <div className="modal-actions">
          <button onClick={onCancel} disabled={applying}>
            {t('paneCancel')}
          </button>
          <button className="primary" onClick={apply} disabled={!loaded || !!error || applying}>
            {applying ? t('paneCutoutApplying') : t('paneCutoutApply')}
          </button>
        </div>
      </div>
    </div>
  )
}
