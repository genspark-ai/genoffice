import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { PageFrame } from '../editor/pagination-gaps'
import type { MultipageLayout } from '../multipage'

/** one canvas element to re-paint on the page it intersects (unzoomed px, relative to the page wrap) */
interface CanvasBlock {
  el: HTMLElement
  top: number
  left: number
  width: number
  height: number
}

export interface CanvasSnapshot {
  wrapW: number
  blocks: CanvasBlock[]
}

/**
 * Every top-level canvas element: the editor's blocks and page-gap widgets plus
 * the wrap-level layers around them (edge header/footer strips, watermark,
 * per-page overlays). Zero-height hosts (float anchors, overlay layers) keep
 * their position: their absolutely placed children carry the paint.
 */
export function snapshotCanvas(wrap: HTMLElement, zoom: number): CanvasSnapshot | null {
  const pm = wrap.querySelector<HTMLElement>('.ProseMirror')
  if (!pm || zoom <= 0) return null
  const wr = wrap.getBoundingClientRect()
  if (wr.width === 0) return null
  const blocks: CanvasBlock[] = []
  const take = (el: HTMLElement) => {
    const r = el.getBoundingClientRect()
    blocks.push({
      el,
      top: (r.top - wr.top) / zoom,
      left: (r.left - wr.left) / zoom,
      width: r.width / zoom,
      height: r.height / zoom,
    })
  }
  for (const el of Array.from(wrap.children) as HTMLElement[]) {
    // the editor root sits inside EditorContent's mount div
    if (el === pm || el.contains(pm)) {
      for (const child of Array.from(pm.children) as HTMLElement[]) take(child)
    } else if (el.tagName !== 'STYLE' && el.tagName !== 'SCRIPT') {
      take(el)
    }
  }
  return { wrapW: wr.width / zoom, blocks }
}

function paintPage(root: HTMLElement, snap: CanvasSnapshot, frame: PageFrame): void {
  root.replaceChildren()
  for (const b of snap.blocks) {
    const inside =
      b.height > 0
        ? b.top < frame.bottom && b.top + b.height > frame.top
        : b.top >= frame.top && b.top < frame.bottom
    if (!inside) continue
    const slot = document.createElement('div')
    slot.className = 'mp-block'
    slot.style.top = `${b.top - frame.top}px`
    slot.style.left = `${b.left}px`
    slot.style.width = `${b.width}px`
    slot.style.height = `${b.height}px`
    const clone = b.el.cloneNode(true) as HTMLElement
    // the slot already sits where the element's margins / offsets put it; a
    // positioned clone stays the containing block of its own absolute children
    clone.style.margin = '0'
    clone.style.transform = 'none'
    clone.style.position = 'relative'
    clone.style.top = '0'
    clone.style.left = '0'
    clone.style.right = 'auto'
    clone.style.bottom = 'auto'
    clone.removeAttribute('id')
    clone.removeAttribute('contenteditable')
    for (const n of clone.querySelectorAll('[contenteditable], [id], [tabindex]')) {
      n.removeAttribute('contenteditable')
      n.removeAttribute('id')
      n.removeAttribute('tabindex')
    }
    slot.appendChild(clone)
    root.appendChild(slot)
  }
}

function PageSlot({
  index,
  frame,
  slot,
  snap,
  scroller,
  onPick,
}: {
  index: number
  frame: PageFrame
  slot: { left: number; top: number; width: number; height: number }
  snap: CanvasSnapshot | null
  scroller: HTMLElement | null
  onPick: (index: number, clientX: number, clientY: number) => void
}) {
  const hostRef = useRef<HTMLDivElement>(null)
  const pageRef = useRef<HTMLDivElement>(null)
  const [visible, setVisible] = useState(false)

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const io = new IntersectionObserver(
      (entries) => setVisible(entries.some((e) => e.isIntersecting)),
      { root: scroller, rootMargin: '50% 0px' },
    )
    io.observe(host)
    return () => io.disconnect()
  }, [scroller])

  useLayoutEffect(() => {
    const root = pageRef.current
    if (!root || !visible || !snap) return
    paintPage(root, snap, frame)
    return () => root.replaceChildren()
  }, [snap, frame, visible])

  return (
    <div
      ref={hostRef}
      className="mp-slot"
      data-page={index + 1}
      style={{ left: slot.left, top: slot.top, width: slot.width, height: slot.height }}
      onMouseDown={(e) => {
        if (e.button !== 0) return
        e.preventDefault()
        e.stopPropagation()
        onPick(index, e.clientX, e.clientY)
      }}
      onClick={(e) => e.stopPropagation()}
    >
      {visible && (
        <div className="page-wrap mp-wrap" style={{ width: snap?.wrapW }}>
          <div
            ref={pageRef}
            className="doc-page ProseMirror mp-page"
            style={{
              clipPath: `inset(0 0 ${Math.max(0, slot.height - (frame.bottom - frame.top))}px 0)`,
            }}
            aria-hidden
          />
        </div>
      )}
    </div>
  )
}

/**
 * The non-active pages of the Multiple Pages grid: each slot re-paints its page
 * from a snapshot of the live canvas (clones placed where the canvas placed
 * them). The active page is not here — the live editor sits over its slot.
 */
export function MultipageGrid({
  frames,
  layout,
  active,
  snap,
  scroller,
  onPick,
}: {
  frames: PageFrame[]
  layout: MultipageLayout
  active: number
  snap: CanvasSnapshot | null
  scroller: HTMLElement | null
  onPick: (index: number, clientX: number, clientY: number) => void
}) {
  return (
    <>
      {layout.slots.map((slot, i) =>
        i === active || !frames[i] ? null : (
          <PageSlot
            key={i}
            index={i}
            frame={frames[i]}
            slot={slot}
            snap={snap}
            scroller={scroller}
            onPick={onPick}
          />
        ),
      )}
    </>
  )
}
