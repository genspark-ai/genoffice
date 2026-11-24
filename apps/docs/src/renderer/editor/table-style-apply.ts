import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command, Transaction } from '@tiptap/pm/state'
import { TableMap, isInTable, selectedRect } from '@tiptap/pm/tables'
import type {
  CellBorder,
  CellBorders,
  TableCondFormat,
  TableLook,
  TableStyleDisplay,
} from '@genoffice/docx-engine'
import { DEFAULT_TABLE_LOOK } from './table-properties'

type Side = 'top' | 'right' | 'bottom' | 'left'

/**
 * What a table style gives one cell: its fill and the lines around it, under
 * the options of the table's w:tblLook. The same precedence the import uses
 * (parse.ts applyTableStyleDisplay): rows beat columns, all beat bands and the
 * whole table; table-level lines go on the outer or inner sides by position.
 */
export function styledCell(
  display: TableStyleDisplay | undefined,
  look: TableLook,
  at: { row: number; rows: number; col: number; span: number; cols: number },
): { fill: string | null; borders: CellBorders | null } {
  if (!display) return { fill: null, borders: null }
  const firstRow = look.firstRow && at.row === 0
  const lastRow = look.lastRow && at.row === at.rows - 1
  const bandSize = display.rowBandSize ?? 1
  const bandRow = look.firstRow ? at.row - 1 : at.row
  const band1 = look.bandedRows && bandRow >= 0 && Math.floor(bandRow / bandSize) % 2 === 0
  const band2 = look.bandedRows && bandRow >= 0 && !band1
  const conds: Array<TableCondFormat | undefined> = [
    firstRow ? display.firstRow : undefined,
    lastRow ? display.lastRow : undefined,
    look.firstColumn && at.col === 0 ? display.firstCol : undefined,
    look.lastColumn && at.col + at.span === at.cols ? display.lastCol : undefined,
    band1 ? display.band1Horz : band2 ? display.band2Horz : undefined,
  ]
  const bandFill = band1 ? display.band1Fill : band2 ? display.band2Fill : undefined
  const fill = conds.find((c) => c?.fill)?.fill ?? bandFill ?? display.fill ?? null

  const table = display.borders ?? {}
  const sides: Partial<Record<Side, CellBorder | undefined>> = {
    top: at.row === 0 ? table.top : table.insideH,
    bottom: at.row === at.rows - 1 ? table.bottom : table.insideH,
    left: at.col === 0 ? table.left : table.insideV,
    right: at.col + at.span === at.cols ? table.right : table.insideV,
  }
  // the first matching condition wins each side
  for (const cond of [...conds].reverse()) if (cond?.borders) Object.assign(sides, cond.borders)
  const borders: CellBorders = {}
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const line = sides[side]
    if (line) borders[side] = line
  }
  return { fill, borders: Object.keys(borders).length > 0 ? borders : null }
}

/**
 * The style's look on each cell of a table, under its w:tblLook, as the
 * display-only styleFill / styleBorders. The cells' own fill and borders are
 * left as they are: they draw over the style, and only they are saved.
 */
function restyleCells(
  tr: Transaction,
  table: PmNode,
  tableStart: number,
  display: TableStyleDisplay | undefined,
  look: TableLook,
  clearOwn: boolean,
): Transaction {
  const map = TableMap.get(table)
  const seen = new Set<number>()
  for (let row = 0; row < map.height; row++) {
    for (let col = 0; col < map.width; col++) {
      const cellPos = map.map[row * map.width + col]
      if (seen.has(cellPos)) continue
      seen.add(cellPos)
      const cell = table.nodeAt(cellPos)
      if (!cell) continue
      const area = map.findCell(cellPos)
      const styled = styledCell(display, look, {
        row: area.top,
        rows: map.height,
        col: area.left,
        span: area.right - area.left,
        cols: map.width,
      })
      tr = tr.setNodeMarkup(tableStart + cellPos, undefined, {
        ...cell.attrs,
        ...(clearOwn ? { fill: null, borders: null } : {}),
        styleFill: styled.fill,
        styleBorders: styled.borders,
      })
    }
  }
  return tr
}

/**
 * Gives the selected table a style, as a click in Word's Table Styles gallery
 * does ("Apply and Clear Formatting"): the table's and its cells' own shading
 * and lines are removed, and the style alone gives the look. Word then saves
 * only w:tblStyle and w:tblLook (probe 2026-11-24: red w:tblBorders, a yellow
 * w:shd and a w:tcBorders were all gone after the style was set), so the look
 * stays in the style: it is drawn from the display-only styleFill /
 * styleBorders and never written into the cells. A table without a w:tblLook
 * gets Word's default one. `display` is the style as read from styles.xml;
 * null clears the style, leaving a table without lines.
 */
export function applyTableStyle(
  styleId: string | null,
  display: TableStyleDisplay | undefined,
): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const { table, tableStart } = selectedRect(state)
    const look = (table.attrs.tblLook as TableLook | null) ?? DEFAULT_TABLE_LOOK
    let tr = restyleCells(state.tr, table, tableStart, styleId ? display : undefined, look, true)
    tr = tr.setNodeMarkup(tableStart - 1, undefined, {
      ...table.attrs,
      tblStyleId: styleId,
      // the table's own lines go too: the save drops its w:tblBorders
      borders: null,
      tblBordersCleared: true,
      // without a w:tblLook Word applies none of the style's conditional formatting
      tblLook: look,
      tblLookEdited: true,
    })
    dispatch?.(tr)
    return true
  }
}

/**
 * After a Table Style Options change: the style's look follows the new
 * w:tblLook, and the cells' own formatting stays untouched.
 */
export function restyleTable(display: TableStyleDisplay | undefined): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const { table, tableStart } = selectedRect(state)
    const look = (table.attrs.tblLook as TableLook | null) ?? DEFAULT_TABLE_LOOK
    dispatch?.(restyleCells(state.tr, table, tableStart, display, look, false))
    return true
  }
}
