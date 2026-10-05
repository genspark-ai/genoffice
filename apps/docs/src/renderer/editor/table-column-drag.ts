import type { EditorView } from '@tiptap/pm/view'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { columnResizingPluginKey } from '@tiptap/pm/tables'
import { MIN_TABLE_COLUMN_PX, setTableGridAtCell, tableGridAtCell } from './table-sizing'

export const tableColumnDragKey = new PluginKey('docTableColumnDrag')

/**
 * Word's border drag: an inner border trades width between the two columns it
 * separates, so the table keeps its width; the last column's right border moves
 * the table edge, up to the section content width. `delta` is in model px.
 */
export function dragColumnBorder(
  start: readonly number[],
  col: number,
  delta: number,
  maxTotal: number,
  minimum = MIN_TABLE_COLUMN_PX,
): number[] {
  const widths = start.slice()
  if (col < 0 || col >= widths.length || !Number.isFinite(delta)) return widths
  if (col < widths.length - 1) {
    const pair = widths[col] + widths[col + 1]
    const floor = Math.min(minimum, pair / 2)
    const left = Math.min(pair - floor, Math.max(floor, widths[col] + delta))
    widths[col] = left
    widths[col + 1] = pair - left
    return widths
  }
  const total = widths.reduce((sum, width) => sum + width, 0)
  // a grid already wider than the page may shrink but never grow further
  const room = Math.max(0, maxTotal - total)
  const floor = Math.min(minimum, widths[col])
  widths[col] = Math.min(widths[col] + room, Math.max(floor, widths[col] + delta))
  return widths
}

function tableDomAt(view: EditorView, tableStart: number): HTMLTableElement | null {
  let dom: Node | null = view.domAtPos(tableStart).node
  while (dom && dom.nodeName !== 'TABLE') dom = dom.parentNode
  return dom as HTMLTableElement | null
}

function sectionContentWidth(view: EditorView, fallback: number): number {
  const raw = getComputedStyle(view.dom).getPropertyValue('--section-content-w')
  const width = Number.parseFloat(raw)
  return Number.isFinite(width) && width > 0 ? width : fallback
}

interface Preview {
  table: HTMLTableElement
  colgroup: HTMLElement
  /** inline styles to put back when the drag ends without a commit */
  tableStyle: string | null
  colStyles: (string | null)[]
  createdCols: HTMLElement[]
}

function startPreview(table: HTMLTableElement, columns: number): Preview | null {
  const colgroup = table.firstElementChild as HTMLElement | null
  if (!colgroup || colgroup.tagName !== 'COLGROUP') return null
  const createdCols: HTMLElement[] = []
  while (colgroup.children.length < columns) {
    const col = document.createElement('col')
    colgroup.appendChild(col)
    createdCols.push(col)
  }
  const cols = Array.from(colgroup.children) as HTMLElement[]
  return {
    table,
    colgroup,
    tableStyle: table.getAttribute('style'),
    colStyles: cols.map((col) => col.getAttribute('style')),
    createdCols,
  }
}

/**
 * Paint a grid onto the table the way renderHTML will after the commit
 * (colgroup percentages of a fixed table width), so the preview is the result.
 */
function paintPreview(preview: Preview, widths: number[], tableCssWidth: number): void {
  const total = widths.reduce((sum, width) => sum + width, 0)
  const { table, colgroup } = preview
  table.style.width = `${tableCssWidth}px`
  table.style.maxWidth = 'none'
  table.style.tableLayout = 'fixed'
  const cols = Array.from(colgroup.children) as HTMLElement[]
  widths.forEach((width, index) => {
    if (cols[index]) cols[index].style.width = `${((width / total) * 100).toFixed(4)}%`
  })
}

function restorePreview(preview: Preview): void {
  const { table, colgroup, tableStyle, colStyles, createdCols } = preview
  if (tableStyle === null) table.removeAttribute('style')
  else table.setAttribute('style', tableStyle)
  const cols = Array.from(colgroup.children) as HTMLElement[]
  cols.forEach((col, index) => {
    const style = colStyles[index]
    if (style === null || style === undefined) col.removeAttribute('style')
    else col.setAttribute('style', style)
  })
  for (const col of createdCols) col.remove()
}

/**
 * Starts a Word-style border drag when prosemirror-tables has a live resize
 * handle. Runs before the library's own mousedown so its single-column preview
 * (which let the whole table shrink or spill past the margin, genoffice#1156)
 * never starts; the library still owns hover detection and the handle widget.
 */
function startDrag(view: EditorView, event: MouseEvent): boolean {
  if (!view.editable || event.button !== 0) return false
  const resize = columnResizingPluginKey.getState(view.state) as
    { activeHandle: number; dragging: unknown } | undefined
  if (!resize || resize.activeHandle < 0 || resize.dragging) return false
  const cellPos = resize.activeHandle
  const $cell = view.state.doc.resolve(cellPos)
  const tableStart = $cell.start(-1)
  const table = tableDomAt(view, tableStart)
  if (!table) return false
  const maxWidth = sectionContentWidth(view, table.parentElement?.clientWidth || 624)
  const grid = tableGridAtCell(view.state, cellPos, maxWidth)
  if (!grid) return false
  const started = startPreview(table, grid.widths.length)
  if (!started) return false
  const preview: Preview = started

  const start = grid.widths
  const startTotal = start.reduce((sum, width) => sum + width, 0)
  const rect = table.getBoundingClientRect()
  // screen px per model px: folds in the canvas zoom and any clamp of the
  // rendered width, so the border stays under the pointer at every zoom
  const screenPerModel = rect.width > 0 ? rect.width / startTotal : 1
  // CSS px per model px (CSS zoom leaves offsetWidth unscaled)
  const cssPerModel = table.offsetWidth > 0 ? table.offsetWidth / startTotal : 1
  const rtl = getComputedStyle(table).direction === 'rtl'
  const startX = event.clientX
  const win = view.dom.ownerDocument.defaultView ?? window
  let current = start

  const widthsAt = (clientX: number): number[] => {
    const delta = ((clientX - startX) / screenPerModel) * (rtl ? -1 : 1)
    return dragColumnBorder(start, grid.col, delta, Math.max(maxWidth, startTotal))
  }
  const paint = (widths: number[]) => {
    const total = widths.reduce((sum, width) => sum + width, 0)
    paintPreview(preview, widths, total * cssPerModel)
  }

  const cleanup = () => {
    win.removeEventListener('mousemove', move)
    win.removeEventListener('mouseup', finish)
    win.removeEventListener('keydown', cancel, true)
    view.dispatch(view.state.tr.setMeta(columnResizingPluginKey, { setDragging: null }))
  }
  function move(moveEvent: MouseEvent) {
    // the button came up outside the window
    if (moveEvent.buttons === 0) return finish(moveEvent)
    current = widthsAt(moveEvent.clientX)
    paint(current)
  }
  function finish(upEvent: MouseEvent) {
    current = widthsAt(upEvent.clientX)
    cleanup()
    const changed = current.some((width, index) => Math.abs(width - start[index]) > 0.01)
    // the commit re-renders the table from the model (same grid as the preview)
    if (!changed || !setTableGridAtCell(cellPos, current)(view.state, view.dispatch)) {
      restorePreview(preview)
    }
  }
  function cancel(keyEvent: KeyboardEvent) {
    if (keyEvent.key !== 'Escape') return
    keyEvent.preventDefault()
    keyEvent.stopPropagation()
    cleanup()
    restorePreview(preview)
  }

  // freeze the handle on this border while dragging (prosemirror-tables stops
  // re-targeting hover and keeps the handle widget) — and the table in place
  view.dispatch(
    view.state.tr.setMeta(columnResizingPluginKey, {
      setDragging: { startX, startWidth: start[grid.col] },
    }),
  )
  paint(start)
  win.addEventListener('mousemove', move)
  win.addEventListener('mouseup', finish)
  win.addEventListener('keydown', cancel, true)
  event.preventDefault()
  return true
}

export function tableColumnDrag(): Plugin {
  return new Plugin({
    key: tableColumnDragKey,
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => startDrag(view, event),
      },
    },
  })
}
