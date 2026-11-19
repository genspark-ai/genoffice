import { Extension } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { TableMap } from '@tiptap/pm/tables'
import { t } from '../i18n/locale'
import { paintPreview, restorePreview, scaleTableGrid, startPreview } from './table-column-drag'
import { setTableGridAtCell, tableGridAtCell } from './table-sizing'
import { previewRowHeights, rowChromePx } from './table-row-drag'

/**
 * Word-style table move handle: a small ⣿-cross button floating at the top-left
 * corner of the hovered table. Clicking it selects the whole table (NodeSelection);
 * dragging it hands the table to ProseMirror's native drag machinery, so dropping
 * repositions the table in the document. A second handle at the bottom-right
 * corner scales the column grid and the row heights, like Word's table resize handle.
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

/** DOM hit-testing remains reliable for CSS-floated tables where posAtCoords may resolve to wrapped text. */
function tablePosFromTarget(view: EditorView, target: EventTarget | null): number | null {
  const element = target instanceof Element ? target : null
  const table = element?.closest('table.doc-table')
  if (!(table instanceof HTMLElement) || !view.dom.contains(table)) return null
  let found: number | null = null
  view.state.doc.descendants((node, pos) => {
    if (found !== null) return false
    if (node.type.name === 'docTable' && view.nodeDOM(pos) === table) {
      found = pos
      return false
    }
    return true
  })
  return found
}

const MOVE_SVG =
  '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" ' +
  'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M8 1v14M1 8h14M8 1 6 3M8 1l2 2M8 15l-2-2M8 15l2-2M1 8l2-2M1 8l2 2M15 8l-2-2M15 8l-2 2"/>' +
  '</svg>'

const RESIZE_SVG =
  '<svg width="10" height="10" viewBox="0 0 16 16" fill="none" stroke="currentColor" ' +
  'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M2 14 14 2M14 7v7H7"/></svg>'

function sectionContentWidth(view: EditorView, fallback: number): number {
  const width = Number.parseFloat(
    getComputedStyle(view.dom).getPropertyValue('--section-content-w'),
  )
  return Number.isFinite(width) && width > 0 ? width : fallback
}

const MIN_ROW_PX = 4

interface RowStart {
  pos: number
  cssHeight: number
}

/** the table's model rows with their painted <tr> (paginated clones are not rows) */
function rowStarts(view: EditorView, tablePos: number): RowStart[] {
  const table = view.state.doc.nodeAt(tablePos)
  if (!table) return []
  const rows: RowStart[] = []
  table.forEach((row, offset) => {
    if (row.type.name !== 'docTableRow') return
    const pos = tablePos + 1 + offset
    const tr = view.nodeDOM(pos)
    if (!(tr instanceof HTMLTableRowElement)) return
    const cssHeight = tr.offsetHeight
    if (cssHeight > 0) rows.push({ pos, cssHeight })
  })
  return rows
}

/** every row keeps its share of the table height; the shortest row stops at the minimum */
export function scaleRowHeights(start: readonly number[], delta: number): number[] {
  const total = start.reduce((sum, height) => sum + height, 0)
  if (!(total > 0) || !Number.isFinite(delta)) return start.slice()
  const minFactor = Math.max(...start.map((height) => Math.min(MIN_ROW_PX, height) / height))
  const factor = Math.max(minFactor, (total + delta) / total)
  return start.map((height) => height * factor)
}

function tableHandlePlugin(): Plugin {
  return new Plugin({
    key: new PluginKey('tableHandle'),
    view(view) {
      const handle = document.createElement('button')
      handle.type = 'button'
      handle.className = 'doc-table-handle'
      handle.style.display = 'none'
      handle.draggable = true
      handle.innerHTML = MOVE_SVG
      handle.title = t('ribbonTableHandleTip')
      handle.setAttribute('aria-label', t('ribbonTableHandleTip'))
      handle.setAttribute('contenteditable', 'false')

      const resizer = document.createElement('button')
      resizer.type = 'button'
      resizer.className = 'doc-table-resize'
      resizer.style.display = 'none'
      resizer.innerHTML = RESIZE_SVG
      resizer.title = t('ribbonTableResizeTip')
      resizer.setAttribute('aria-label', t('ribbonTableResizeTip'))
      resizer.setAttribute('contenteditable', 'false')

      // tiptap mounts the view into a temporary element and React later moves its
      // children (these buttons included) under the real one, so the host is
      // re-resolved on every placement instead of captured once
      const host = (): HTMLElement | null => {
        const parent = view.dom.parentElement
        if (!parent) return null
        parent.style.position ||= 'relative'
        if (handle.parentElement !== parent) parent.append(handle, resizer)
        return parent
      }
      host()

      let tablePos: number | null = null
      let suppressClick = false
      let stopFloatDrag: (() => void) | null = null

      const hideHandle = () => {
        handle.style.display = 'none'
        resizer.style.display = 'none'
        tablePos = null
      }

      const placeHandle = (pos: number) => {
        const container = host()
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
        placeResizer(dom)
      }

      const placeResizer = (dom: HTMLElement) => {
        const container = host()
        if (!container) return
        const tableRect = dom.getBoundingClientRect()
        const containerRect = container.getBoundingClientRect()
        const scale = container.offsetWidth ? containerRect.width / container.offsetWidth : 1
        resizer.style.display = 'flex'
        resizer.style.top = `${(tableRect.bottom - containerRect.top) / scale + container.scrollTop + 2}px`
        resizer.style.left = `${(tableRect.right - containerRect.left) / scale + 2}px`
      }

      const onMouseMove = (event: MouseEvent) => {
        if (!view.editable) return
        const pos =
          tablePosFromTarget(view, event.target) ??
          tablePosAt(view, { left: event.clientX, top: event.clientY })
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
        if (suppressClick) {
          suppressClick = false
          return
        }
        const selection = selectTable()
        if (!selection) return
        view.dispatch(view.state.tr.setSelection(selection))
        view.focus()
      }

      const onDragStart = (event: DragEvent) => {
        if (!event.dataTransfer) return
        const selection = selectTable()
        if (!selection) return event.preventDefault()
        if (selection.node.attrs.tblFloat === 'left' || selection.node.attrs.tblFloat === 'right') {
          event.preventDefault()
          return
        }
        view.dispatch(view.state.tr.setSelection(selection))
        view.dragging = { slice: selection.content(), move: true }
        event.dataTransfer.effectAllowed = 'move'
        event.dataTransfer.setData('text/plain', ' ')
        const dom = view.nodeDOM(selection.from)
        if (dom instanceof HTMLElement) event.dataTransfer.setDragImage(dom, 0, 0)
      }

      // A floating table uses the same drag-to-position interaction as Word's
      // anchored pictures. In-flow tables keep ProseMirror's native block move.
      const onMouseDown = (event: MouseEvent) => {
        if (event.button !== 0 || tablePos === null) return
        const dragPos = tablePos
        const node = view.state.doc.nodeAt(dragPos)
        if (!node || (node.attrs.tblFloat !== 'left' && node.attrs.tblFloat !== 'right')) return
        event.preventDefault()
        event.stopPropagation()

        const visual = view.nodeDOM(dragPos)
        if (!(visual instanceof HTMLElement)) return
        const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
        const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
        const startX = event.clientX
        const startY = event.clientY
        const initialX = Number(node.attrs.tblFloatXTwips) || 0
        const initialY = Number(node.attrs.tblFloatYTwips) || 0
        const priorTransform = visual.style.transform

        const cleanup = () => {
          window.removeEventListener('mousemove', onMove)
          window.removeEventListener('mouseup', onUp)
          stopFloatDrag = null
        }
        const delta = (e: MouseEvent) => ({
          x: (e.clientX - startX) / zoom,
          y: (e.clientY - startY) / zoom,
        })
        const onMove = (e: MouseEvent) => {
          const { x, y } = delta(e)
          visual.style.transform = `translate(${x.toFixed(1)}px,${y.toFixed(1)}px)`
        }
        const onUp = (e: MouseEvent) => {
          cleanup()
          visual.style.transform = priorTransform
          const { x, y } = delta(e)
          if (Math.hypot(x, y) < 3) return
          suppressClick = true
          window.setTimeout(() => {
            suppressClick = false
          }, 0)
          const current = view.state.doc.nodeAt(dragPos)
          if (!current || current.type.name !== 'docTable') return
          view.dispatch(
            view.state.tr.setNodeMarkup(dragPos, undefined, {
              ...current.attrs,
              tblFloatSource: current.attrs.tblFloat,
              tblFloatSuppressed: false,
              tblFloatXTwips: Math.round(initialX + x * 15),
              tblFloatYTwips: Math.round(initialY + y * 15),
              tblFloatEdited: true,
            }),
          )
        }
        stopFloatDrag?.()
        stopFloatDrag = cleanup
        window.addEventListener('mousemove', onMove)
        window.addEventListener('mouseup', onUp)
      }

      // Word's corner handle: the pointer's horizontal travel scales every
      // column by the same factor; the preview paints the grid the commit will
      // render, so what the user sees while dragging is the result.
      const onResizeDown = (event: MouseEvent) => {
        if (event.button !== 0 || tablePos === null || !view.editable) return
        const node = view.state.doc.nodeAt(tablePos)
        if (!node || node.type.name !== 'docTable') return
        const tableStart = tablePos + 1
        const cellPos = tableStart + TableMap.get(node).map[0]
        const table = view.nodeDOM(tablePos)
        if (!(table instanceof HTMLTableElement)) return
        const maxWidth = sectionContentWidth(view, table.parentElement?.clientWidth || 624)
        const grid = tableGridAtCell(view.state, cellPos, maxWidth)
        if (!grid) return
        const preview = startPreview(table, grid.widths.length)
        if (!preview) return
        event.preventDefault()
        event.stopPropagation()
        cancelHide()

        const start = grid.widths
        const startTotal = start.reduce((sum, width) => sum + width, 0)
        const rect = table.getBoundingClientRect()
        const screenPerModel = rect.width > 0 ? rect.width / startTotal : 1
        const cssPerModel = table.offsetWidth > 0 ? table.offsetWidth / startTotal : 1
        const rows = rowStarts(view, tablePos)
        const rowHeights = rows.map((row) => row.cssHeight)
        const screenPerCss =
          rect.height > 0 && table.offsetHeight > 0 ? rect.height / table.offsetHeight : 1
        // the handle sits on the visual right edge whatever the text direction
        const startX = event.clientX
        const startY = event.clientY
        const widthsAt = (clientX: number) =>
          scaleTableGrid(start, (clientX - startX) / screenPerModel, maxWidth)
        const heightsAt = (clientY: number) =>
          scaleRowHeights(rowHeights, (clientY - startY) / screenPerCss)
        const paint = (widths: number[], heights: number[]) => {
          const total = widths.reduce((sum, width) => sum + width, 0)
          // the row decoration redraws the rows, which leaves the colgroup alone
          if (rows.length) {
            previewRowHeights(
              view,
              rows.map((row, index) => ({ pos: row.pos, heightPx: heights[index] })),
            )
          }
          paintPreview(preview, widths, total * cssPerModel)
          placeResizer(table)
        }
        const restore = () => {
          restorePreview(preview)
          if (rows.length) previewRowHeights(view, null)
        }
        const cleanup = () => {
          window.removeEventListener('mousemove', onMove)
          window.removeEventListener('mouseup', onUp)
          window.removeEventListener('keydown', onKey, true)
          stopFloatDrag = null
        }
        const onMove = (e: MouseEvent) => {
          if (e.buttons === 0) return onUp(e)
          paint(widthsAt(e.clientX), heightsAt(e.clientY))
        }
        const onUp = (e: MouseEvent) => {
          cleanup()
          const widths = widthsAt(e.clientX)
          const heights = heightsAt(e.clientY)
          const widthChanged = widths.some((width, index) => Math.abs(width - start[index]) > 0.01)
          const heightChanged = heights.some(
            (height, index) => Math.abs(height - rowHeights[index]) > 0.5,
          )
          restore()
          hideHandle()
          if (!widthChanged && !heightChanged) return
          let tr: Transaction | null = null
          if (widthChanged) setTableGridAtCell(cellPos, widths)(view.state, (next) => (tr = next))
          tr ??= view.state.tr
          commitRowHeights(tr, rows, heightChanged ? heights : null)
        }
        const onKey = (e: KeyboardEvent) => {
          if (e.key !== 'Escape') return
          e.preventDefault()
          e.stopPropagation()
          cleanup()
          restore()
          hideHandle()
        }
        // w:trHeight is a content floor under the cell margins and gridlines, so
        // the painted row is measured and the chrome taken back out (same undo step)
        const commitRowHeights = (tr: Transaction, rows: RowStart[], heights: number[] | null) => {
          const twipsOf = (row: RowStart, cssHeight: number) => ({
            ...view.state.doc.nodeAt(row.pos)!.attrs,
            heightTwips: Math.round(Math.max(MIN_ROW_PX, cssHeight) * 15),
            heightRule:
              view.state.doc.nodeAt(row.pos)!.attrs.heightRule === 'exact' ? 'exact' : 'atLeast',
          })
          if (heights) {
            rows.forEach((row, index) =>
              tr.setNodeMarkup(row.pos, undefined, twipsOf(row, heights[index])),
            )
          }
          view.dispatch(tr)
          if (!heights) return
          let fix: Transaction | null = null
          rows.forEach((row, index) => {
            const painted = view.nodeDOM(row.pos)
            const attrs = view.state.doc.nodeAt(row.pos)?.attrs
            if (!(painted instanceof HTMLElement) || !attrs) return
            const chrome = rowChromePx(painted, attrs.heightTwips)
            if (chrome <= 0.5) return
            fix ??= view.state.tr.setMeta('addToHistory', false)
            fix.setNodeMarkup(row.pos, undefined, {
              ...attrs,
              heightTwips: Math.max(15, Math.round((heights[index] - chrome) * 15)),
            })
          })
          if (fix) view.dispatch(fix)
        }
        stopFloatDrag?.()
        stopFloatDrag = () => {
          cleanup()
          restore()
        }
        paint(start, rowHeights)
        window.addEventListener('mousemove', onMove)
        window.addEventListener('mouseup', onUp)
        window.addEventListener('keydown', onKey, true)
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
      handle.addEventListener('mousedown', onMouseDown)
      handle.addEventListener('click', onClick)
      handle.addEventListener('dragstart', onDragStart)
      resizer.addEventListener('mousemove', (e) => e.stopPropagation())
      resizer.addEventListener('mouseenter', cancelHide)
      resizer.addEventListener('mouseleave', scheduleHide)
      resizer.addEventListener('mousedown', onResizeDown)
      document.addEventListener('scroll', onScroll, true)

      return {
        update(_view, prevState) {
          // any doc change invalidates the captured position; the next hover re-syncs
          if (!prevState.doc.eq(view.state.doc)) hideHandle()
        },
        destroy() {
          stopFloatDrag?.()
          cancelHide()
          view.dom.removeEventListener('mousemove', onMouseMove)
          view.dom.removeEventListener('mouseleave', scheduleHide)
          document.removeEventListener('scroll', onScroll, true)
          handle.remove()
          resizer.remove()
        },
      }
    },
  })
}
