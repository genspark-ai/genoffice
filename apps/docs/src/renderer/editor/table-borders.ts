import { CellSelection, isInTable, selectedRect, type TableRect } from '@tiptap/pm/tables'
import type { Command } from '@tiptap/pm/state'

/** Word's Borders ▾ gallery entries in ribbon order */
export const TABLE_BORDER_MENU = [
  'bottom',
  'top',
  'left',
  'right',
  'none',
  'all',
  'outer',
  'inner',
  'insideH',
  'insideV',
] as const

export type TableBorderMode = (typeof TABLE_BORDER_MENU)[number]

export interface BorderLine {
  style: string
  szEighths?: number
  color?: string
}

type Sides = Record<string, BorderLine>
const SIDES = ['top', 'bottom', 'left', 'right'] as const
type Side = (typeof SIDES)[number]

const NONE: BorderLine = { style: 'none' }

/** Word's "Apply to" rule: selected cells get the borders; a bare caret formats the whole table */
function targetRect(rect: TableRect, cellSelection: boolean): TableRect {
  if (cellSelection) return rect
  return { ...rect, top: 0, left: 0, bottom: rect.map.height, right: rect.map.width }
}

export function sideLine(
  mode: TableBorderMode,
  side: Side,
  edge: Record<Side, boolean>,
  line: BorderLine,
): BorderLine | undefined {
  const horizontal = side === 'top' || side === 'bottom'
  switch (mode) {
    case 'all':
      return line
    case 'none':
      return NONE
    case 'outer':
      return edge[side] ? line : undefined
    case 'inner':
      return edge[side] ? undefined : line
    case 'insideH':
      return horizontal && !edge[side] ? line : undefined
    case 'insideV':
      return !horizontal && !edge[side] ? line : undefined
    default:
      return mode === side && edge[side] ? line : undefined
  }
}

/** Applies one Borders ▾ entry with the current pen (style / weight / color) to the
 *  selected cells, or to the whole table when nothing is selected. */
export function setSelectionBorders(mode: TableBorderMode, line: BorderLine): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = targetRect(selectedRect(state), state.selection instanceof CellSelection)
    let tr = state.tr
    const seen = new Set<number>()
    for (let row = rect.top; row < rect.bottom; row++) {
      for (let col = rect.left; col < rect.right; col++) {
        const cellPos = rect.map.map[row * rect.map.width + col]
        if (seen.has(cellPos)) continue
        seen.add(cellPos)
        const pos = rect.tableStart + cellPos
        const node = state.doc.nodeAt(pos)
        if (!node) continue
        const cellRect = rect.map.findCell(cellPos)
        const edge = {
          top: cellRect.top <= rect.top,
          bottom: cellRect.bottom >= rect.bottom,
          left: cellRect.left <= rect.left,
          right: cellRect.right >= rect.right,
        }
        const next: Sides = { ...((node.attrs.borders as Sides | null) ?? {}) }
        let changed = false
        for (const side of SIDES) {
          const value = sideLine(mode, side, edge, line)
          if (value) {
            next[side] = value
            changed = true
          }
        }
        if (changed) tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders: next })
      }
    }
    if (mode === 'none') {
      // table-level inside lines would otherwise survive No Border
      const tablePos = rect.tableStart - 1
      const tableNode = tr.doc.nodeAt(tablePos)
      const prev = tableNode?.attrs.borders as Sides | null | undefined
      if (tableNode?.type.name === 'docTable' && prev && (prev.insideH || prev.insideV)) {
        const { insideH: _h, insideV: _v, ...rest } = prev
        tr = tr.setNodeMarkup(tablePos, undefined, {
          ...tableNode.attrs,
          borders: Object.keys(rest).length ? rest : null,
        })
      }
    }
    if (!tr.docChanged) return false
    dispatch?.(tr)
    return true
  }
}
