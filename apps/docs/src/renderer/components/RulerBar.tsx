import { useRef, useState, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useRulerFollow } from './ruler-follow'

/** height of the horizontal ruler bar for a zoom factor, px */
export const rulerBarHeight = (zoom: number): number => Math.ceil(22 * zoom) + 6

/**
 * Word's horizontal ruler bar: fixed above the document view, outside the
 * scrolling pages. The ruler inside is drawn at the document zoom and lined up
 * with the page (centered over the page column, following horizontal scroll).
 */
export function RulerBar({
  editor,
  zoom,
  pageWidthPx,
  children,
}: {
  editor: Editor | null
  /** document zoom factor (1 = 100%) */
  zoom: number
  /** the ruler's unzoomed width: the page width */
  pageWidthPx: number
  children: ReactNode
}) {
  const barRef = useRef<HTMLDivElement>(null)
  const [left, setLeft] = useState<number | null>(null)
  useRulerFollow(editor, () => {
    const bar = barRef.current
    if (!bar || !editor || editor.isDestroyed) return
    const live = (editor.view.dom as HTMLElement).closest('.doc-live')
    if (!live) return
    const lr = live.getBoundingClientRect()
    const br = bar.getBoundingClientRect()
    const next = (lr.left + (lr.width - pageWidthPx * zoom) / 2 - br.left) / zoom
    setLeft((prev) => (prev !== null && Math.abs(prev - next) < 0.25 ? prev : next))
  })
  return (
    <div className="ruler-bar" ref={barRef} style={{ height: rulerBarHeight(zoom) }}>
      {left !== null && (
        <div className="ruler-bar-host" style={{ zoom, left }}>
          {children}
        </div>
      )}
    </div>
  )
}
