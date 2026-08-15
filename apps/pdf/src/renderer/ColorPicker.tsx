import { useLayoutEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { useI18n } from './i18n/locale'
import {
  COLOR_PRESETS,
  hexTo255,
  hsvToRgb,
  isHexColor,
  rgb255ToHex,
  rgbToHsv,
} from './color-presets'

interface ColorPickerPopoverProps {
  /** Current color as #rrggbb — seeds the gradient and marks the matching swatch */
  value?: string
  /** Quick-pick swatches; defaults to the app-wide shared palette */
  presets?: readonly string[]
  /** Extra classes on the popover root (e.g. `rb-drop` when anchored to a ribbon trigger) */
  className?: string
  onPick: (hex: string) => void
  onClose: () => void
}

interface EyeDropperApi {
  open: () => Promise<{ sRGBHex: string }>
}

const clamp01 = (n: number): number => Math.min(1, Math.max(0, n))

/** The single color picker used across the PDF app: saturation/value gradient +
    hue slider (+ eyedropper when Chromium exposes one), quick preset swatches
    that apply on click, and a free hex field. OK confirms the gradient/hex
    color; callers own the open state and outside-click behavior. */
export function ColorPickerPopover({
  value,
  presets = COLOR_PRESETS,
  className,
  onPick,
  onClose,
}: ColorPickerPopoverProps): ReactElement {
  const { t } = useI18n()
  const [hsv, setHsv] = useState<[number, number, number]>(() =>
    value && isHexColor(value) ? rgbToHsv(...hexTo255(value)) : [0, 0, 0],
  )
  const hex = rgb255ToHex(hsvToRgb(...hsv))
  const [hexDraft, setHexDraft] = useState(hex.toUpperCase())
  const dragging = useRef<'sv' | 'hue' | null>(null)
  const svRef = useRef<HTMLDivElement>(null)
  const hueRef = useRef<HTMLDivElement>(null)
  const rootRef = useRef<HTMLDivElement>(null)

  // Anchored/floating parents (ribbon anchor positioning, the text-edit
  // editor's page coordinates) land the popover on fractional pixels, which
  // smears every 1px hairline inside — snap it to whole device pixels
  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el) return
    el.style.transform = ''
    const dpr = window.devicePixelRatio || 1
    const r = el.getBoundingClientRect()
    const dx = Math.round(r.left * dpr) / dpr - r.left
    const dy = Math.round(r.top * dpr) / dpr - r.top
    if (dx || dy) el.style.transform = `translate(${dx}px, ${dy}px)`
  }, [])

  const setColor = (next: [number, number, number]): void => {
    setHsv(next)
    setHexDraft(rgb255ToHex(hsvToRgb(...next)).toUpperCase())
  }
  const pick = (h: string): void => {
    onPick(h.toLowerCase())
    onClose()
  }

  const svFromEvent = (e: ReactPointerEvent): void => {
    const r = svRef.current?.getBoundingClientRect()
    if (!r) return
    setColor([
      hsv[0],
      clamp01((e.clientX - r.left) / r.width),
      clamp01(1 - (e.clientY - r.top) / r.height),
    ])
  }
  const hueFromEvent = (e: ReactPointerEvent): void => {
    const r = hueRef.current?.getBoundingClientRect()
    if (!r) return
    setColor([clamp01((e.clientX - r.left) / r.width) * 359.9, hsv[1], hsv[2]])
  }

  const eyeDropper = (window as { EyeDropper?: new () => EyeDropperApi }).EyeDropper

  return (
    <div ref={rootRef} className={`pdf-color-popover${className ? ` ${className}` : ''}`}>
      <div
        ref={svRef}
        className="pdf-cp-sv"
        style={{
          background: `linear-gradient(to top, #000, rgba(0, 0, 0, 0)),
            linear-gradient(to right, #fff, hsl(${hsv[0]} 100% 50%))`,
        }}
        onPointerDown={(e) => {
          e.preventDefault()
          e.currentTarget.setPointerCapture(e.pointerId)
          dragging.current = 'sv'
          svFromEvent(e)
        }}
        onPointerMove={(e) => dragging.current === 'sv' && svFromEvent(e)}
        onPointerUp={() => (dragging.current = null)}
      >
        <span
          className="pdf-cp-thumb"
          style={{
            left: `${hsv[1] * 100}%`,
            top: `${(1 - hsv[2]) * 100}%`,
            background: hex,
            border: '2px solid #fff',
            boxShadow: '0 0 0 1px rgb(0 0 0 / 35%)',
          }}
        />
      </div>
      <div className="pdf-cp-row">
        {eyeDropper && (
          <button
            type="button"
            className="pdf-cp-eyedrop"
            aria-label={t('moreColors')}
            onClick={() => {
              new eyeDropper()
                .open()
                .then((res) => {
                  if (isHexColor(res.sRGBHex)) setColor(rgbToHsv(...hexTo255(res.sRGBHex)))
                })
                .catch(() => {
                  /* eyedropper dismissed */
                })
            }}
          >
            <svg
              viewBox="0 0 24 24"
              width="15"
              height="15"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M19.4 4.6a2.1 2.1 0 0 0-3 0L14 7l-1-1-1.8 1.8 5 5L18 11l-1-1 2.4-2.4a2.1 2.1 0 0 0 0-3Z" />
              <path d="m12.5 8.5-7 7c-.4.4-.6.9-.7 1.4l-.4 2.3a.6.6 0 0 0 .4.4l2.3-.4c.5-.1 1-.3 1.4-.7l7-7" />
            </svg>
          </button>
        )}
        <span className="pdf-cp-preview" style={{ background: hex }} />
        <div
          ref={hueRef}
          className="pdf-cp-hue"
          style={{
            background: 'linear-gradient(to right, #f00, #ff0, #0f0, #0ff, #00f, #f0f, #f00)',
          }}
          onPointerDown={(e) => {
            e.preventDefault()
            e.currentTarget.setPointerCapture(e.pointerId)
            dragging.current = 'hue'
            hueFromEvent(e)
          }}
          onPointerMove={(e) => dragging.current === 'hue' && hueFromEvent(e)}
          onPointerUp={() => (dragging.current = null)}
        >
          <span
            className="pdf-cp-thumb"
            style={{
              left: `${(hsv[0] / 360) * 100}%`,
              top: '50%',
              background: `hsl(${hsv[0]} 100% 50%)`,
              border: '2px solid #fff',
              boxShadow: '0 0 0 1px rgb(0 0 0 / 35%)',
            }}
          />
        </div>
      </div>
      {/* Integer-pixel tracks (6×31 + 5×6 gaps = exactly the 216px content width):
          fractional 1fr tracks put the 1px swatch borders on half pixels and blur them */}
      <div className="pdf-color-palette" style={{ gridTemplateColumns: 'repeat(6, 31px)', gap: 6 }}>
        {presets.map((p) => (
          <button
            key={p}
            type="button"
            className={`pdf-color-choice${hex.toUpperCase() === p.toUpperCase() ? ' active' : ''}`}
            style={{ background: p }}
            aria-label={p}
            onClick={() => pick(p)}
          />
        ))}
      </div>
      <input
        className="pdf-modal-input pdf-color-hex"
        value={hexDraft}
        maxLength={7}
        placeholder="#000000"
        aria-label={t('moreColors')}
        // Some ancestors (ribbon markup group) preventDefault on mousedown to
        // keep the document text selection — the hex field still needs focus
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          const v = e.target.value
          setHexDraft(v)
          if (isHexColor(v)) setHsv(rgbToHsv(...hexTo255(v)))
        }}
        onBlur={() => setHexDraft(hex.toUpperCase())}
        onKeyDown={(e) => {
          if (e.key === 'Enter') pick(isHexColor(hexDraft) ? hexDraft : hex)
        }}
      />
      <div className="pdf-color-popover-actions">
        <button type="button" className="pdf-modal-btn" onClick={onClose}>
          {t('cancel')}
        </button>
        <button type="button" className="pdf-modal-btn primary" onClick={() => pick(hex)}>
          {t('ok')}
        </button>
      </div>
    </div>
  )
}
