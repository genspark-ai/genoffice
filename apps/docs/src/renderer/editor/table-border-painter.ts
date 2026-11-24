import { Extension } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { TableMap } from '@tiptap/pm/tables'
import type { CellBorder, CellBorders } from '@genoffice/docx-engine'

/**
 * Word's Border Painter (Table Design ▸ Borders): while it is on, a click on a
 * cell draws the current pen (line style, weight, color) on the cell edge
 * nearest the pointer, and dragging keeps drawing on each edge passed. The
 * neighbour across the edge gets the same line, so the collapsed border reads
 * the same from both cells. Escape, or the button again, turns it off.
 */

type Side = 'top' | 'right' | 'bottom' | 'left'

const OPPOSITE: Record<Side, Side> = { top: 'bottom', bottom: 'top', left: 'right', right: 'left' }

export interface TableBorderPainterStorage {
  active: boolean
  pen: CellBorder
  /** told when the painter turns on or off, so the ribbon button can follow */
  onChange: ((active: boolean) => void) | null
}

declare module '@tiptap/core' {
  interface Storage {
    tableBorderPainter: TableBorderPainterStorage
  }
}

/** the edge of a cell box nearest a point */
export function nearestSide(
  box: { left: number; top: number; right: number; bottom: number },
  x: number,
  y: number,
): Side {
  const distances: Array<[Side, number]> = [
    ['top', Math.abs(y - box.top)],
    ['bottom', Math.abs(box.bottom - y)],
    ['left', Math.abs(x - box.left)],
    ['right', Math.abs(box.right - x)],
  ]
  return distances.reduce((best, next) => (next[1] < best[1] ? next : best))[0]
}

/** the cell's own pos and its table around a DOM cell, or null outside an editable body table */
function cellAt(view: EditorView, el: Element | null) {
  const dom = el?.closest('td, th')
  if (!dom || !view.dom.contains(dom)) return null
  let pos: number
  try {
    pos = view.posAtDOM(dom, 0) - 1
  } catch {
    return null
  }
  const $cell = view.state.doc.resolve(pos)
  const cell = $cell.nodeAfter
  if (!cell || cell.type.spec.tableRole?.includes('cell') !== true) return null
  const table = $cell.node(-1)
  if (table.type.spec.tableRole !== 'table') return null
  return { dom, pos, cell, table, tableStart: $cell.start(-1) }
}

/** draws the pen on one side of a cell and on the neighbour sharing that edge */
export function paintEdge(view: EditorView, cellPos: number, side: Side, pen: CellBorder): void {
  const { state } = view
  const $cell = state.doc.resolve(cellPos)
  const table = $cell.node(-1)
  const tableStart = $cell.start(-1)
  const map = TableMap.get(table)
  const rect = map.findCell(cellPos - tableStart)
  const setSide = (pos: number, which: Side) => {
    const node = tr.doc.nodeAt(pos)
    if (!node) return
    const borders = { ...((node.attrs.borders as CellBorders | null) ?? {}), [which]: pen }
    tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders })
  }
  let tr = state.tr
  setSide(cellPos, side)
  // the cells across the edge, along the whole edge of a merged cell
  const across: number[] = []
  if (side === 'top' && rect.top > 0)
    for (let c = rect.left; c < rect.right; c++) across.push((rect.top - 1) * map.width + c)
  if (side === 'bottom' && rect.bottom < map.height)
    for (let c = rect.left; c < rect.right; c++) across.push(rect.bottom * map.width + c)
  if (side === 'left' && rect.left > 0)
    for (let r = rect.top; r < rect.bottom; r++) across.push(r * map.width + rect.left - 1)
  if (side === 'right' && rect.right < map.width)
    for (let r = rect.top; r < rect.bottom; r++) across.push(r * map.width + rect.right)
  for (const pos of new Set(across.map((index) => map.map[index])))
    setSide(tableStart + pos, OPPOSITE[side])
  view.dispatch(tr)
}

export const TableBorderPainterExtension = Extension.create<object, TableBorderPainterStorage>({
  name: 'tableBorderPainter',
  addStorage() {
    return {
      active: false,
      pen: { style: 'single', szEighths: 4, color: '000000' },
      onChange: null,
    }
  },
  addProseMirrorPlugins() {
    const storage = this.storage
    let painting = false
    let last = ''
    const paintAt = (view: EditorView, event: MouseEvent) => {
      const hit = cellAt(view, event.target as Element | null)
      if (!hit) return false
      const box = hit.dom.getBoundingClientRect()
      const side = nearestSide(box, event.clientX, event.clientY)
      const key = `${hit.pos}:${side}`
      if (key !== last) {
        last = key
        paintEdge(view, hit.pos, side, storage.pen)
      }
      return true
    }
    return [
      new Plugin({
        props: {
          attributes: (): Record<string, string> =>
            storage.active ? { class: 'table-border-painting' } : {},
          handleKeyDown(_view, event) {
            if (!storage.active || event.key !== 'Escape') return false
            storage.active = false
            storage.onChange?.(false)
            _view.dispatch(_view.state.tr)
            return true
          },
          handleDOMEvents: {
            mousedown(view, event) {
              if (!storage.active || event.button !== 0 || !view.editable) return false
              last = ''
              if (!paintAt(view, event)) return false
              painting = true
              event.preventDefault()
              const up = () => {
                painting = false
                window.removeEventListener('mouseup', up)
              }
              window.addEventListener('mouseup', up)
              return true
            },
            mousemove(view, event) {
              if (!storage.active || !painting) return false
              paintAt(view, event)
              return true
            },
          },
        },
      }),
    ]
  },
})
