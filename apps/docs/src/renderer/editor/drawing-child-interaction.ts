// Modified for ForaOffice: existing object gestures address the pressed drawing child.
import { Plugin } from '@tiptap/pm/state'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { selectBoxAt } from './selected-box'
import { drawingLineSize } from './drawing-line'

export function drawingChildInteraction(): Plugin {
  let cancelActive: (() => void) | null = null
  return new Plugin({
    view: () => ({
      update: (view, previous) => {
        if (view.state.doc !== previous.doc || !view.editable) cancelActive?.()
      },
      destroy: () => cancelActive?.(),
    }),
    props: {
      handleDOMEvents: {
        mousedown(view, event) {
          if (event.button !== 0 || event.detail >= 2 || !view.editable) return false
          const target = event.target as HTMLElement
          const rotator = target.closest<HTMLElement>('.box-rotator')
          const el = rotator?.querySelector<HTMLElement>('.doc-textbox')
          const wrapper = rotator?.closest<HTMLElement>('.doc-protected')
          if (!rotator || !el || !wrapper || wrapper.classList.contains('doc-content-editing'))
            return false
          let pos = -1
          view.state.doc.descendants((node, at) => {
            if (pos < 0 && node.type.name === 'docProtected' && view.nodeDOM(at) === wrapper)
              pos = at
            return pos < 0
          })
          const node = pos >= 0 ? view.state.doc.nodeAt(pos) : null
          const boxes = node?.attrs.textboxes as TextboxDisplay[] | null
          const index = Array.from(wrapper.querySelectorAll('.doc-textbox')).indexOf(el)
          if (!boxes || index < 0 || !boxes[index].drawingSource?.grouped) return false
          cancelActive?.()
          const box = boxes[index]
          const resize = target.classList.contains('box-resize-handle')
          const rotate = target.classList.contains('box-rotate-handle')
          event.preventDefault()
          event.stopPropagation()
          view.dispatch(selectBoxAt(view.state.tr, pos, index))
          view.focus()
          const gestureDoc = view.state.doc
          const zoomEl = wrapper.closest<HTMLElement>('.doc-zoom')
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom) || 1 : 1
          const x = event.clientX
          const y = event.clientY
          const rect = el.getBoundingClientRect()
          const angle = (e: MouseEvent) =>
            (Math.atan2(
              e.clientY - rect.top - rect.height / 2,
              e.clientX - rect.left - rect.width / 2,
            ) *
              180) /
            Math.PI
          const startAngle = angle(event)
          const transform = rotator.style.transform
          const originalWidth = el.style.width
          const originalHeight = el.style.height
          let changed = false
          let patch: Partial<TextboxDisplay> = {}
          const onMove = (e: MouseEvent) => {
            if (view.state.doc !== gestureDoc || !view.editable) {
              cancelActive?.()
              return
            }
            const dx = (e.clientX - x) / zoom
            const dy = (e.clientY - y) / zoom
            if (!changed && Math.hypot(dx, dy) < 3) return
            changed = true
            if (rotate) {
              let deg = ((Math.round((box.rotDeg ?? 0) + angle(e) - startAngle) % 360) + 360) % 360
              if (e.shiftKey) deg = (Math.round(deg / 15) * 15) % 360
              patch = { rotDeg: deg }
              rotator.style.transform = `rotate(${deg}deg)`
            } else if (resize) {
              const rad = ((box.rotDeg ?? 0) * Math.PI) / 180
              const localX = dx * Math.cos(rad) + dy * Math.sin(rad)
              const localY = -dx * Math.sin(rad) + dy * Math.cos(rad)
              const w = box.widthPx === 0 ? 0 : Math.max(8, (box.widthPx ?? 189) + localX)
              const h = box.heightPx === 0 ? 0 : Math.max(8, (box.heightPx ?? 113) + localY)
              patch = { widthPx: w, heightPx: h, minHeightPx: h }
              const size = box.lineStyle ? drawingLineSize({ ...box, ...patch }) : { w, h }
              el.style.width = `${Math.max(8, size.w)}px`
              el.style.height = `${Math.max(8, size.h)}px`
            } else {
              patch = {
                offsetXEmu: (box.offsetXEmu ?? 0) + Math.round(dx * 9525),
                offsetYEmu: (box.offsetYEmu ?? 0) + Math.round(dy * 9525),
              }
              rotator.style.transform = `translate(${dx}px,${dy}px) ${transform}`
            }
          }
          const onUp = (e: MouseEvent) => {
            if (view.state.doc !== gestureDoc || !view.editable) {
              cancelActive?.()
              return
            }
            onMove(e)
            cleanup()
            if (!changed) return
            const current = view.state.doc.nodeAt(pos)
            const next = current?.attrs.textboxes as TextboxDisplay[] | undefined
            if (!current || !next) return
            view.dispatch(
              selectBoxAt(
                view.state.tr.setNodeMarkup(pos, undefined, {
                  ...current.attrs,
                  textboxes: next.map((b, i) => (i === index ? { ...b, ...patch } : b)),
                }),
                pos,
                index,
              ),
            )
          }
          const cleanup = () => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
            window.removeEventListener('blur', cleanup)
            rotator.style.transform = transform
            if (view.state.doc === gestureDoc) {
              el.style.width = originalWidth
              el.style.height = originalHeight
            }
            cancelActive = null
          }
          cancelActive = cleanup
          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          window.addEventListener('blur', cleanup)
          return true
        },
      },
    },
  })
}
