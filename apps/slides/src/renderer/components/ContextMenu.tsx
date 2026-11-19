/**
 * Context menu — fixed positioning, kept within the viewport; closes on any click/Escape or a
 * scroll outside the menu. A menu taller than the viewport scrolls its own body instead.
 * Items that are null render as separators; items with `sub` open a one-level flyout on hover.
 */
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useEscOverlay } from '../esc-overlay'
import { platformShortcuts } from '@genoffice/i18n'

export interface CtxItem {
  label: string
  /** Shortcut hint on the right */
  hint?: string
  disabled?: boolean
  danger?: boolean
  /** Check mark on the left (toggle items) */
  checked?: boolean
  onClick?: () => void
  /** Flyout items (one level deep; sub-items can't nest further) */
  sub?: Array<CtxItem | null>
  /** Row of color swatches rendered under the label instead of a clickable item ('none' = the clear square) */
  swatches?: string[]
  onSwatch?: (color: string) => void
}

interface Props {
  x: number
  y: number
  items: Array<CtxItem | null>
  onClose: () => void
  /** Menu over a live text edit: presses must not move focus (the overlay commits on blur) */
  keepEdit?: boolean
}

export function ContextMenu({ x, y, items, onClose, keepEdit }: Props) {
  useEscOverlay(true)
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ x, y })
  const [openSub, setOpenSub] = useState<number | null>(null)
  const subRef = useRef<HTMLDivElement>(null)
  const subAnchor = useRef<DOMRect | null>(null)
  const [subStyle, setSubStyle] = useState<React.CSSProperties>()

  const openSubAt = (i: number, host: Element) => {
    subAnchor.current = host.getBoundingClientRect()
    setOpenSub(i)
  }

  // The flyout is fixed-positioned from the host row's rect: the menu body scrolls when
  // taller than the viewport, and an absolutely-positioned flyout would be clipped by
  // that scroll box
  useLayoutEffect(() => {
    const el = subRef.current
    const anchor = subAnchor.current
    if (openSub == null || !el || !anchor) return
    const r = el.getBoundingClientRect()
    const left =
      anchor.right + 2 + r.width > window.innerWidth
        ? Math.max(0, anchor.left - r.width - 2)
        : anchor.right + 2
    const top = Math.max(4, Math.min(anchor.top - 5, window.innerHeight - r.height - 4))
    setSubStyle({ position: 'fixed', left, top, margin: 0 })
  }, [openSub])

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
    // Capture: Escape only closes the menu, it must not reach the text editor (commit) or the canvas (deselect)
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape') return
      e.preventDefault()
      e.stopPropagation()
      onClose()
    }
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    // Wheel inside the menu scrolls an overflowing menu body; only an outside scroll closes
    const onWheel = (e: WheelEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    window.addEventListener('keydown', onKey, true)
    window.addEventListener('mousedown', onDown)
    window.addEventListener('wheel', onWheel)
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('keydown', onKey, true)
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('wheel', onWheel)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  const swatchRow = (swatches: string[], onSwatch?: (color: string) => void) => (
    <div className="ctx-swatches-row">
      {swatches.map((c) => (
        <button
          key={c}
          className={`ctx-swatch ${c === 'none' ? 'ctx-swatch-none' : ''}`}
          style={c === 'none' ? undefined : { background: c }}
          data-tip={c}
          aria-label={c}
          onClick={() => {
            onClose()
            onSwatch?.(c)
          }}
        >
          {c === 'none' ? '✕' : ''}
        </button>
      ))}
    </div>
  )

  const plainItem = (item: CtxItem, key: number) => (
    <button
      key={key}
      className={`ctx-item ${item.danger ? 'danger' : ''}${item.checked ? ' ctx-checked' : ''}`}
      disabled={item.disabled}
      onClick={() => {
        onClose()
        item.onClick?.()
      }}
    >
      <span>{item.label}</span>
      {item.hint && <span className="ctx-hint">{platformShortcuts(item.hint)}</span>}
    </button>
  )

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => e.preventDefault()}
      onScroll={() => setOpenSub(null)}
      {...(keepEdit ? { 'data-keep-edit': '', onMouseDown: (e) => e.preventDefault() } : {})}
    >
      {items.map((item, i) =>
        item === null ? (
          <div key={i} className="ctx-sep" />
        ) : item.swatches ? (
          <div key={i} className="ctx-swatches">
            <span className="ctx-swatches-label">{item.label}</span>
            {swatchRow(item.swatches, item.onSwatch)}
          </div>
        ) : item.sub ? (
          <div
            key={i}
            className="ctx-sub-host"
            onMouseEnter={(e) => !item.disabled && openSubAt(i, e.currentTarget)}
            onMouseLeave={() => setOpenSub(null)}
          >
            <button
              className={`ctx-item${openSub === i ? ' ctx-sub-open' : ''}`}
              disabled={item.disabled}
              onClick={(e) => openSubAt(i, e.currentTarget)}
            >
              <span>{item.label}</span>
              <span className="ctx-hint">▸</span>
            </button>
            {openSub === i && (
              <div ref={subRef} className="ctx-menu ctx-submenu" style={subStyle}>
                {item.sub.map((it, j) =>
                  it === null ? <div key={j} className="ctx-sep" /> : plainItem(it, j),
                )}
              </div>
            )}
          </div>
        ) : (
          plainItem(item, i)
        ),
      )}
    </div>
  )
}
