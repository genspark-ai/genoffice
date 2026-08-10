/**
 * Context menu — fixed positioning, kept within the viewport; closes on any click/Escape/scroll.
 * Items that are null render as separators.
 */
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { platformShortcuts } from '@genoffice/i18n'

export interface CtxItem {
  label: string
  /** Shortcut hint on the right */
  hint?: string
  disabled?: boolean
  danger?: boolean
  onClick?: () => void
  /** Row of color swatches rendered under the label instead of a clickable item ('none' = the clear square) */
  swatches?: string[]
  onSwatch?: (color: string) => void
}

interface Props {
  x: number
  y: number
  items: Array<CtxItem | null>
  onClose: () => void
}

export function ContextMenu({ x, y, items, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ x, y })

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const { innerWidth, innerHeight } = window
    const r = el.getBoundingClientRect()
    setPos({
      x: Math.min(x, Math.max(0, innerWidth - r.width - 4)),
      y: Math.min(y, Math.max(0, innerHeight - r.height - 4)),
    })
  }, [x, y])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('mousedown', onDown)
    window.addEventListener('wheel', onClose, { once: true })
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('wheel', onClose)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {items.map((item, i) =>
        item === null ? (
          <div key={i} className="ctx-sep" />
        ) : item.swatches ? (
          <div key={i} className="ctx-swatches">
            <span className="ctx-swatches-label">{item.label}</span>
            <div className="ctx-swatches-row">
              {item.swatches.map((c) => (
                <button
                  key={c}
                  className={`ctx-swatch ${c === 'none' ? 'ctx-swatch-none' : ''}`}
                  style={c === 'none' ? undefined : { background: c }}
                  data-tip={c}
                  aria-label={c}
                  onClick={() => {
                    onClose()
                    item.onSwatch?.(c)
                  }}
                >
                  {c === 'none' ? '✕' : ''}
                </button>
              ))}
            </div>
          </div>
        ) : (
          <button
            key={i}
            className={`ctx-item ${item.danger ? 'danger' : ''}`}
            disabled={item.disabled}
            onClick={() => {
              onClose()
              item.onClick?.()
            }}
          >
            <span>{item.label}</span>
            {item.hint && <span className="ctx-hint">{platformShortcuts(item.hint)}</span>}
          </button>
        ),
      )}
    </div>
  )
}
