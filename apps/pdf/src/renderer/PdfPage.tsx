import { useEffect, useRef, useState } from 'react'
import type { CSSProperties, RefObject } from 'react'
import { AnnotationMode, TextLayer } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import { pdfRectToCss, quadToRect } from './annotations'
import type { LocalMarkup, PageGeom } from './annotations'

/** Which items in the container are within the (expanded) viewport — shared lazy-render basis
    for pages/thumbnails. Rebuild the observer when enabled flips (sidebar toggles unmount/remount the root) */
export function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
  enabled = true,
): { visible: Set<number>; setItemRef: (idx: number) => (el: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])
  useEffect(() => {
    const root = rootRef.current
    if (!enabled || !root || count === 0) return
    const io = new IntersectionObserver(
      (entries) => {
        setVisible((prev) => {
          const next = new Set(prev)
          for (const e of entries) {
            const idx = Number((e.target as HTMLElement).dataset.idx)
            if (e.isIntersecting) next.add(idx)
            else next.delete(idx)
          }
          return next
        })
      },
      { root, rootMargin },
    )
    for (const el of itemRefs.current) if (el) io.observe(el)
    return () => io.disconnect()
  }, [rootRef, count, rootMargin, enabled])
  return {
    visible,
    setItemRef: (idx) => (el) => {
      itemRefs.current[idx] = el
    },
  }
}

/** Single page: renders canvas + text layer (select/copy) when visible, released once off-viewport */
export function PdfPage({
  doc,
  pageNo,
  scale,
  rotationDelta,
  visible,
  onRenderState,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  /** Unsaved rotation delta (clockwise degrees) */
  rotationDelta: number
  visible: boolean
  onRenderState: (doc: PDFDocumentProxy, pageNo: number, pending: boolean) => void
}) {
  const holderRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const holder = holderRef.current
    if (!holder) return
    // Offscreen pages still release their bitmap. For an in-place rerender (save
    // reload, zoom, rotation), keep the previous canvas visible until its replacement
    // is fully rendered so the page never flashes white between document instances.
    if (!visible) {
      holder.replaceChildren()
      // A page captured by the post-save barrier may scroll out before its replacement
      // renders. Its overlays are no longer mounted, so treat the cleared offscreen page
      // as settled instead of making the whole document wait for the timeout.
      onRenderState(doc, pageNo, false)
      return
    }
    // Visibility may change while a save reload is running. Register newly visible
    // pages dynamically so global preview cleanup cannot outrun their bitmap swap.
    onRenderState(doc, pageNo, true)
    let cancelled = false
    let renderTask: RenderTask | null = null
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const viewport = page.getViewport({ scale, rotation: (page.rotate + rotationDelta) % 360 })
      // Cap at 2x: on hi-dpi screens a 3x-dpr full-page bitmap doubles memory with no visible gain
      const dpr = Math.min(window.devicePixelRatio || 1, 2)
      const canvas = document.createElement('canvas')
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        // FormLayer renders interactive widgets as HTML. Exclude their saved appearance
        // streams from the page bitmap or filled values are drawn twice after a reload.
        annotationMode: AnnotationMode.ENABLE_FORMS,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        return // cancelled
      }
      if (cancelled) return
      const textDiv = document.createElement('div')
      textDiv.className = 'textLayer'
      holder.replaceChildren(canvas, textDiv)
      // Notify after the bitmap swap, but before the browser paints. A post-save
      // reload uses this to remove the matching edit previews in the same frame.
      onRenderState(doc, pageNo, false)
      const textLayer = new TextLayer({
        textContentSource: page.streamTextContent(),
        container: textDiv,
        viewport,
      })
      try {
        await textLayer.render()
      } catch {
        /* cancelled */
      }
    })()
    return () => {
      cancelled = true
      renderTask?.cancel()
    }
  }, [doc, pageNo, scale, rotationDelta, visible, onRenderState])
  return <div ref={holderRef} className="pdf-page-content" />
}

/** Overlay for unsaved markups. Purely visual (pointer-events: none) so the text
 *  underneath stays selectable; clicking is handled by the page-level hit test. */
export function MarkupOverlay({
  markups,
  geom,
  scale,
  selectedId,
}: {
  markups: LocalMarkup[]
  geom: PageGeom
  scale: number
  selectedId: string | null
}) {
  return (
    <>
      {markups.flatMap((m) =>
        m.quads.map((q, i) => {
          const [r, g, b] = m.color
          const style: CSSProperties = pdfRectToCss(geom, quadToRect(q), scale)
          if (m.type === 'highlight') {
            style.background = `rgba(${r * 255}, ${g * 255}, ${b * 255}, 0.4)`
          } else {
            const bar = `rgb(${r * 255}, ${g * 255}, ${b * 255})`
            if (m.type === 'underline') style.borderBottom = `2px solid ${bar}`
            else style.backgroundImage = `linear-gradient(${bar}, ${bar})`
          }
          return (
            <div
              key={`${m.id}-${i}`}
              className={`pdf-markup pdf-markup-${m.type}${m.id === selectedId ? ' pdf-markup-selected' : ''}`}
              style={style}
            />
          )
        }),
      )}
    </>
  )
}
