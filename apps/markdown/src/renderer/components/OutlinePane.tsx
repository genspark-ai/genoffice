import { useRef, useState } from 'react'
import type { PointerEvent, ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { OutlineItem } from '../editor/outline'

export const OUTLINE_MIN_WIDTH = 160
export const OUTLINE_MAX_WIDTH = 480
const OUTLINE_DEFAULT_WIDTH = 232

export function clampOutlineWidth(width: number): number {
  return Math.min(OUTLINE_MAX_WIDTH, Math.max(OUTLINE_MIN_WIDTH, Math.round(width)))
}

/**
 * Document outline sidebar: heading list in document order, indented by level
 * (PDF bookmark parity). Clicking an entry jumps the editor cursor there; the
 * right edge drags to resize.
 */
export function OutlinePane({
  items,
  onJump,
  width,
  onResize,
}: {
  items: OutlineItem[]
  onJump: (pos: number) => void
  width?: number
  onResize: (width: number) => void
}): ReactElement {
  const { t } = useI18n()
  const [dragging, setDragging] = useState(false)
  const drag = useRef<{ x: number; width: number } | null>(null)
  const current = clampOutlineWidth(width ?? OUTLINE_DEFAULT_WIDTH)

  const onPointerDown = (e: PointerEvent<HTMLDivElement>) => {
    if (e.button !== 0) return
    drag.current = { x: e.clientX, width: current }
    e.currentTarget.setPointerCapture(e.pointerId)
    setDragging(true)
  }
  const onPointerMove = (e: PointerEvent<HTMLDivElement>) => {
    if (!drag.current) return
    onResize(clampOutlineWidth(drag.current.width + e.clientX - drag.current.x))
  }
  const onPointerUp = () => {
    drag.current = null
    setDragging(false)
  }

  return (
    <>
      <aside className="md-outline" aria-label={t('outline')} style={{ width: current }}>
        <div className="md-outline-title">{t('outline')}</div>
        {items.length === 0 ? (
          <div className="md-outline-empty">{t('outlineEmpty')}</div>
        ) : (
          <nav className="md-outline-list">
            {items.map((item, index) => (
              <button
                key={`${item.pos}-${index}`}
                type="button"
                className="md-outline-item"
                style={{ paddingLeft: 10 + (item.level - 1) * 14 }}
                data-tip={item.text}
                title={item.text}
                onClick={() => onJump(item.pos)}
              >
                {item.text}
              </button>
            ))}
          </nav>
        )}
      </aside>
      <div
        className={`md-outline-resizer${dragging ? ' dragging' : ''}`}
        role="separator"
        aria-orientation="vertical"
        aria-label={t('outlineResize')}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
      />
    </>
  )
}
