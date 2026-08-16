import { Extension } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { t } from '../i18n/locale'

/**
 * Word-style table move handle: a small ⣿-cross button floating at the top-left
 * corner of the hovered table. Clicking it selects the whole table (NodeSelection);
 * dragging it hands the table to ProseMirror's native drag machinery, so dropping
 * repositions the table in the document.
 */
export const TableHandle = Extension.create({
  name: 'tableHandle',

  addProseMirrorPlugins() {
    return [tableHandlePlugin()]
  },
})

/** absolute pos of the docTable ancestor under the given viewport coords (null = not in a table) */
function tablePosAt(view: EditorView, coords: { left: number; top: number }): number | null {
  const found = view.posAtCoords(coords)
  if (!found) return null
  const $pos = view.state.doc.resolve(found.inside >= 0 ? found.inside : found.pos)
  for (let depth = 1; depth <= $pos.depth; depth++) {
    if ($pos.node(depth).type.name === 'docTable') return $pos.before(depth)
  }
  if (found.inside >= 0 && view.state.doc.nodeAt(found.inside)?.type.name === 'docTable') {
    return found.inside
  }
  return null
}

const MOVE_SVG =
  '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" ' +
  'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M8 1v14M1 8h14M8 1 6 3M8 1l2 2M8 15l-2-2M8 15l2-2M1 8l2-2M1 8l2 2M15 8l-2-2M15 8l-2 2"/>' +
  '</svg>'

function tableHandlePlugin(): Plugin {
  return new Plugin({
    key: new PluginKey('tableHandle'),
    view(view) {
      const container = view.dom.parentElement

      const handle = document.createElement('button')
      handle.type = 'button'
      handle.className = 'doc-table-handle'
      handle.style.display = 'none'
      handle.draggable = true
      handle.innerHTML = MOVE_SVG
      handle.title = t('ribbonTableHandleTip')
      handle.setAttribute('aria-label', t('ribbonTableHandleTip'))
      handle.setAttribute('contenteditable', 'false')

      if (container) {
        container.style.position ||= 'relative'
        container.append(handle)
      }

      let tablePos: number | null = null

      const hideHandle = () => {
        handle.style.display = 'none'
        tablePos = null
      }

      const placeHandle = (pos: number) => {
        if (!container) return
        const dom = view.nodeDOM(pos)
        if (!(dom instanceof HTMLElement)) return hideHandle()
        const tableRect = dom.getBoundingClientRect()
        const containerRect = container.getBoundingClientRect()
        // DOMRect values are viewport pixels, absolute offsets inside the zoomed
        // page are layout pixels — divide the deltas so the handle scales once
        // with the page instead of twice
        const scale = container.offsetWidth ? containerRect.width / container.offsetWidth : 1
        tablePos = pos
        handle.style.display = 'flex'
        handle.style.top = `${(tableRect.top - containerRect.top) / scale + container.scrollTop - 18}px`
        handle.style.left = `${(tableRect.left - containerRect.left) / scale - 18}px`
      }

      const onMouseMove = (event: MouseEvent) => {
        if (!view.editable) return
        const pos = tablePosAt(view, { left: event.clientX, top: event.clientY })
        if (pos === null) {
          // moving toward the handle itself must not hide it (it sits outside the table box)
          if (tablePos !== null) scheduleHide()
          return
        }
        cancelHide()
        if (pos !== tablePos) placeHandle(pos)
      }

      /** tablePos was captured on a past mousemove — validate against the current doc */
      const selectTable = (): NodeSelection | null => {
        if (tablePos === null || tablePos >= view.state.doc.content.size) return null
        if (view.state.doc.nodeAt(tablePos)?.type.name !== 'docTable') return null
        try {
          return NodeSelection.create(view.state.doc, tablePos)
        } catch {
          return null
        }
      }

      const onClick = () => {
        const selection = selectTable()
        if (!selection) return
        view.dispatch(view.state.tr.setSelection(selection))
        view.focus()
      }

      const onDragStart = (event: DragEvent) => {
        if (!event.dataTransfer) return
        const selection = selectTable()
        if (!selection) return event.preventDefault()
        view.dispatch(view.state.tr.setSelection(selection))
        view.dragging = { slice: selection.content(), move: true }
        event.dataTransfer.effectAllowed = 'move'
        event.dataTransfer.setData('text/plain', ' ')
        const dom = view.nodeDOM(selection.from)
        if (dom instanceof HTMLElement) event.dataTransfer.setDragImage(dom, 0, 0)
      }

      // the handle floats outside the table's corner, so reaching it always
      // leaves the table box first — hide on a grace period, not instantly
      let hideTimer: number | null = null
      const cancelHide = () => {
        if (hideTimer !== null) window.clearTimeout(hideTimer)
        hideTimer = null
      }
      const scheduleHide = () => {
        cancelHide()
        hideTimer = window.setTimeout(() => {
          hideTimer = null
          hideHandle()
        }, 300)
      }
      const onScroll = () => {
        cancelHide()
        hideHandle()
      }

      view.dom.addEventListener('mousemove', onMouseMove)
      view.dom.addEventListener('mouseleave', scheduleHide)
      handle.addEventListener('mousemove', (e) => e.stopPropagation())
      handle.addEventListener('mouseenter', cancelHide)
      handle.addEventListener('mouseleave', scheduleHide)
      handle.addEventListener('click', onClick)
      handle.addEventListener('dragstart', onDragStart)
      document.addEventListener('scroll', onScroll, true)

      return {
        update(_view, prevState) {
          // any doc change invalidates the captured position; the next hover re-syncs
          if (!prevState.doc.eq(view.state.doc)) hideHandle()
        },
        destroy() {
          cancelHide()
          view.dom.removeEventListener('mousemove', onMouseMove)
          view.dom.removeEventListener('mouseleave', scheduleHide)
          document.removeEventListener('scroll', onScroll, true)
          handle.remove()
        },
      }
    },
  })
}
