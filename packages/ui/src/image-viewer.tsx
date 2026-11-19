/**
 * Full-window picture viewer shared by the editors: fit / actual size / wheel
 * zoom, drag to pan, optional Save. Apps pass their own strings through
 * `labels` and import `image-viewer.css` for the chrome.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent, ReactNode, WheelEvent as ReactWheelEvent } from 'react'
import type { Lang } from '@genoffice/i18n'
import { useImageViewerFocus } from './image-viewer-focus'
import { IMAGE_VIEWER_TITLES } from './strings-image-viewer'

/** marks the control the viewer focuses on open and restores focus to after Esc */
const CLOSE_CONTROL = '[data-imgview-close]'

export interface ImageViewerLabels {
  zoomIn: string
  zoomOut: string
  actualSize: string
  fitToWindow: string
  save: string
  close: string
}

const MIN_SCALE = 0.05
const MAX_SCALE = 16
const STEP = 1.25
/** breathing room around a fitted picture (px per side) */
const FIT_PAD = 48
const BAR_HEIGHT = 56

function Svg({ children }: { children: ReactNode }) {
  return (
    <svg
      width={16}
      height={16}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function ImageViewer({
  src,
  alt,
  labels,
  onClose,
  onSave,
  lang = 'en',
}: {
  src: string
  alt?: string
  labels: ImageViewerLabels
  onClose: () => void
  onSave?: () => void
  /** language of the dialog's accessible name */
  lang?: Lang
}) {
  const [natural, setNatural] = useState<{ w: number; h: number } | null>(null)
  const [scale, setScale] = useState(1)
  const [fit, setFit] = useState(true)
  const [offset, setOffset] = useState({ x: 0, y: 0 })
  const drag = useRef<{ x: number; y: number; ox: number; oy: number } | null>(null)
  const [dragging, setDragging] = useState(false)
  const imgRef = useRef<HTMLImageElement>(null)

  // a cached (data:) picture can be complete before React attaches onLoad
  const measure = useCallback(() => {
    const el = imgRef.current
    if (el?.complete && el.naturalWidth) setNatural({ w: el.naturalWidth, h: el.naturalHeight })
  }, [])
  useEffect(() => {
    setNatural(null)
    measure()
  }, [src, measure])

  const fitScale = useCallback(() => {
    if (!natural) return 1
    const w = window.innerWidth - FIT_PAD * 2
    const h = window.innerHeight - FIT_PAD * 2 - BAR_HEIGHT
    return Math.min(1, w / natural.w, h / natural.h)
  }, [natural])

  useEffect(() => {
    if (!fit) return
    setScale(fitScale())
    setOffset({ x: 0, y: 0 })
    const onResize = () => setScale(fitScale())
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [fit, fitScale])

  const clamp = (v: number) => Math.min(MAX_SCALE, Math.max(MIN_SCALE, v))
  const zoomTo = useCallback((next: number) => {
    setFit(false)
    setScale(clamp(next))
  }, [])
  // functional update: rapid wheel / key-repeat ticks in one frame must compound
  const zoomBy = useCallback((factor: number) => {
    setFit(false)
    setScale((s) => clamp(s * factor))
  }, [])

  // Esc closes through useImageViewerFocus on the mask; the zoom keys stay on
  // window so they work wherever focus sits inside the viewer, and the editor
  // behind the overlay must not see them either.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === '+' || e.key === '=') zoomBy(STEP)
      else if (e.key === '-') zoomBy(1 / STEP)
      else if (e.key === '0') {
        setOffset({ x: 0, y: 0 })
        zoomTo(1)
      } else return
      e.preventDefault()
      e.stopPropagation()
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [zoomBy, zoomTo])

  // Esc / Tab trap / initial focus on the close control / focus restore (one handler).
  const focus = useImageViewerFocus(onClose, CLOSE_CONTROL)

  const onWheel = (e: ReactWheelEvent) => {
    e.preventDefault()
    zoomBy(Math.exp(-e.deltaY * 0.002))
  }

  const onMouseDown = (e: ReactMouseEvent) => {
    if (e.button !== 0) return
    e.preventDefault()
    drag.current = { x: e.clientX, y: e.clientY, ox: offset.x, oy: offset.y }
    setDragging(true)
  }
  useEffect(() => {
    if (!dragging) return
    const onMove = (e: MouseEvent) => {
      const d = drag.current
      if (d) setOffset({ x: d.ox + e.clientX - d.x, y: d.oy + e.clientY - d.y })
    }
    const onUp = () => {
      drag.current = null
      setDragging(false)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [dragging])

  const button = (
    label: string,
    onClick: () => void,
    icon: ReactNode,
    active = false,
    extra: { 'data-imgview-close'?: true } = {},
  ) => (
    <button
      type="button"
      className={`gs-imgview-btn${active ? ' active' : ''}`}
      aria-label={label}
      title={label}
      onClick={onClick}
      {...extra}
    >
      {icon}
    </button>
  )

  return (
    <div
      className="gs-imgview-mask"
      role="dialog"
      aria-modal="true"
      aria-label={IMAGE_VIEWER_TITLES[lang]}
      tabIndex={-1}
      ref={focus.ref}
      onKeyDown={focus.onKeyDown}
      onBlur={focus.onBlur}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
      onWheel={onWheel}
    >
      <div className="gs-imgview-bar" onClick={(e) => e.stopPropagation()}>
        {button(
          labels.zoomOut,
          () => zoomBy(1 / STEP),
          <Svg>
            <circle cx="7" cy="7" r="4.5" />
            <path d="M10.5 10.5L14 14M5 7h4" />
          </Svg>,
        )}
        <span className="gs-imgview-zoom">{Math.round(scale * 100)}%</span>
        {button(
          labels.zoomIn,
          () => zoomBy(STEP),
          <Svg>
            <circle cx="7" cy="7" r="4.5" />
            <path d="M10.5 10.5L14 14M5 7h4M7 5v4" />
          </Svg>,
        )}
        <span className="gs-imgview-sep" />
        {button(
          labels.actualSize,
          () => {
            setOffset({ x: 0, y: 0 })
            zoomTo(1)
          },
          <span className="gs-imgview-11">1:1</span>,
          !fit && scale === 1,
        )}
        {button(
          labels.fitToWindow,
          () => setFit(true),
          <Svg>
            <path d="M2 6V2h4M10 2h4v4M14 10v4h-4M6 14H2v-4" />
          </Svg>,
          fit,
        )}
        {onSave && (
          <>
            <span className="gs-imgview-sep" />
            {button(
              labels.save,
              onSave,
              <Svg>
                <path d="M8 2v8M4.5 6.5L8 10l3.5-3.5M2.5 12.5v1h11v-1" />
              </Svg>,
            )}
          </>
        )}
        <span className="gs-imgview-sep" />
        {button(
          labels.close,
          onClose,
          <Svg>
            <path d="M3.5 3.5l9 9M12.5 3.5l-9 9" />
          </Svg>,
          false,
          { 'data-imgview-close': true },
        )}
      </div>
      <img
        className={`gs-imgview-img${dragging ? ' dragging' : ''}`}
        src={src}
        alt={alt ?? ''}
        ref={imgRef}
        draggable={false}
        onLoad={measure}
        onMouseDown={onMouseDown}
        onClick={(e) => e.stopPropagation()}
        style={{
          width: natural ? natural.w : undefined,
          height: natural ? natural.h : undefined,
          transform: `translate(${offset.x}px, ${offset.y}px) scale(${scale})`,
          visibility: natural ? 'visible' : 'hidden',
        }}
      />
    </div>
  )
}
