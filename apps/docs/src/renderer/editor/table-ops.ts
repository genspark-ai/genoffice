import type { Fragment, Node as PmNode, Schema } from '@tiptap/pm/model'
import type { Command, EditorState, Transaction } from '@tiptap/pm/state'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import {
  CellSelection,
  TableMap,
  deleteColumn,
  deleteRow,
  isInTable,
  mergeCells,
  selectedRect,
} from '@tiptap/pm/tables'
import { alignAttrFor, effectiveBidi } from './direction'
import { showToast } from '../components/toast-bus'
import { t } from '../i18n/locale'
import {
  columnWidthsPx,
  copyCellFormatting,
  insertColumnAt,
  insertRowAt,
  pctOf,
  reflowColumns,
} from '../ai/table-ops'

/** Word's Table menu operations that prosemirror-tables has no command for.
 *  Structural edits rebuild the whole table node from a placed-cell grid: a
 *  single replace step keeps colspan/rowspan bookkeeping in one place. */

interface GridCell {
  node: PmNode
  top: number
  left: number
  rowspan: number
  colspan: number
}

interface Grid {
  map: TableMap
  cells: GridCell[]
  rows: PmNode[]
}

const clamp = (n: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, Math.floor(n)))

/** split `total` into `parts` near-equal integers, the remainder going to the first ones */
const partition = (total: number, parts: number): number[] =>
  Array.from({ length: parts }, (_, i) => Math.floor(total / parts) + (i < total % parts ? 1 : 0))

function gridOf(table: PmNode): Grid {
  const map = TableMap.get(table)
  const cells: GridCell[] = []
  const seen = new Set<number>()
  for (const pos of map.map) {
    if (seen.has(pos)) continue
    seen.add(pos)
    const r = map.findCell(pos)
    cells.push({
      node: table.nodeAt(pos)!,
      top: r.top,
      left: r.left,
      rowspan: r.bottom - r.top,
      colspan: r.right - r.left,
    })
  }
  const rows: PmNode[] = []
  table.forEach((row) => rows.push(row))
  return { map, cells, rows }
}

function buildTable(
  table: PmNode,
  rows: PmNode[],
  cells: GridCell[],
  attrs: Record<string, unknown> = table.attrs,
): PmNode {
  const byRow: GridCell[][] = rows.map(() => [])
  for (const c of cells) byRow[c.top]?.push(c)
  return table.type.create(
    attrs,
    rows.map((row, i) =>
      row.type.create(
        row.attrs,
        byRow[i]
          .sort((a, b) => a.left - b.left)
          .map((c) =>
            c.node.type.create(
              { ...c.node.attrs, colspan: c.colspan, rowspan: c.rowspan },
              c.node.content,
            ),
          ),
      ),
    ),
  )
}

/** an empty paragraph shaped like the cell's first one (Word keeps the paragraph format) */
function blankCellContent(schema: Schema, like: PmNode): Fragment | PmNode {
  const first = like.firstChild
  const attrs =
    first?.type.name === 'docParagraph' ? { ...first.attrs, docxIndex: null } : { docxIndex: null }
  return schema.nodes.docParagraph.create(attrs)
}

function cellWidths(cell: GridCell): number[] | null {
  const w = cell.node.attrs.colwidth as number[] | null
  return Array.isArray(w) && w.length === cell.colspan && w.every((v) => v > 0) ? w : null
}

const blankRowAttrs = (row: PmNode): Record<string, unknown> => ({
  ...row.attrs,
  heightTwips: null,
  heightRule: null,
  repeatHeader: false,
  repeatHeaderEdited: false,
  // the copied cantSplit has no raw trPr behind it: mark it edited so the save writes it
  cantSplitEdited: true,
  rawTrPr: null,
  rowRevision: null,
})

/**
 * Split the cell covering (row, col) into R×C cells. Extra columns widen the
 * grid inside the cell's span (other rows' cells spanning it grow their
 * colspan and keep their width); extra rows are inserted after the cell's
 * bottom edge (neighbours ending there grow their rowspan).
 */
export function splitGridCell(
  table: PmNode,
  at: { row: number; col: number },
  rowsN: number,
  colsN: number,
): { table: PmNode; focus: { top: number; left: number } } | null {
  const R = clamp(rowsN, 1, 50)
  const C = clamp(colsN, 1, 50)
  const g = gridOf(table)
  const T = g.cells.find(
    (c) =>
      c.top <= at.row &&
      at.row < c.top + c.rowspan &&
      c.left <= at.col &&
      at.col < c.left + c.colspan,
  )
  if (!T || T.node.attrs.gridGap || (R === 1 && C === 1)) return null
  const { colspan: cs, rowspan: rs } = T

  // ---- columns: k[i] new grid columns per old column i of the cell's span
  const extraCols = Math.max(0, C - cs)
  const k = extraCols > 0 ? partition(C, cs) : Array.from({ length: cs }, () => 1)
  const prefix = (i: number) => k.slice(0, i).reduce((a, b) => a + b, 0)
  const newCol = (x: number) =>
    x <= T.left ? x : x < T.left + cs ? T.left + prefix(x - T.left) : x + extraCols
  const remapWidths = (cell: GridCell): number[] | null => {
    const w = cellWidths(cell)
    if (!w) return null
    const out: number[] = []
    for (let x = cell.left; x < cell.left + cell.colspan; x++) {
      const inSpan = x >= T.left && x < T.left + cs
      const kk = inSpan ? k[x - T.left] : 1
      for (let j = 0; j < kk; j++) out.push(w[x - cell.left] / kk)
    }
    return out
  }
  const placed: GridCell[] = g.cells
    .filter((c) => c !== T)
    .map((c) => {
      const left = newCol(c.left)
      const colspan = newCol(c.left + c.colspan) - left
      const colwidth = remapWidths(c)
      const node =
        colwidth || c.node.attrs.colwidth
          ? c.node.type.create({ ...c.node.attrs, colwidth }, c.node.content)
          : c.node
      return { ...c, node, left, colspan }
    })

  // ---- rows
  const extraRows = Math.max(0, R - rs)
  const insertAt = T.top + rs
  const rows = [...g.rows]
  for (let i = 0; i < extraRows; i++) {
    const template = g.rows[insertAt - 1]
    rows.splice(insertAt, 0, template.type.create(blankRowAttrs(template)))
  }
  for (const c of placed) {
    if (c.top >= insertAt) c.top += extraRows
    else if (c.top + c.rowspan >= insertAt) c.rowspan += extraRows
  }

  // ---- the pieces
  const spans = partition(Math.max(cs, C), C)
  const heights = partition(Math.max(rs, R), R)
  const tW = cellWidths(T)
  const regionWidths = tW
    ? k.flatMap((kk, i) => Array.from({ length: kk }, () => tW[i] / kk))
    : null
  const blank = blankCellContent(table.type.schema, T.node)
  let top = T.top
  for (let ri = 0; ri < R; ri++) {
    let left = T.left
    let cursor = 0
    for (let ci = 0; ci < C; ci++) {
      const span = spans[ci]
      const colwidth = regionWidths ? regionWidths.slice(cursor, cursor + span) : null
      const node = T.node.type.create(
        { ...T.node.attrs, colwidth },
        ri === 0 && ci === 0 ? T.node.content : blank,
      )
      placed.push({ node, top, left, rowspan: heights[ri], colspan: span })
      left += span
      cursor += span
    }
    top += heights[ri]
  }

  const attrs = { ...table.attrs }
  const pct = table.attrs.colWidthsPct as number[] | null
  if (pct && pct.length === g.map.width) {
    attrs.colWidthsPct = pct.flatMap((p, x) =>
      x >= T.left && x < T.left + cs
        ? Array.from({ length: k[x - T.left] }, () => p / k[x - T.left])
        : [p],
    )
  }
  return { table: buildTable(table, rows, placed, attrs), focus: { top: T.top, left: T.left } }
}

function caretIntoCell(
  tr: Transaction,
  tablePos: number,
  table: PmNode,
  top: number,
  left: number,
) {
  const map = TableMap.get(table)
  const pos = tablePos + 1 + map.map[top * map.width + left]
  tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1)))
}

/** Split Cells… dialog: the selection's top-left cell (or the merged selection) becomes rows × cols */
export function splitCellsInto(rows: number, cols: number, mergeFirst: boolean): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    const tablePos = rect.tableStart - 1
    let tr = state.tr
    // a caret inside an already merged cell also spans several grid columns: only a
    // real multi-cell selection is merged first
    const sel = state.selection
    if (mergeFirst && sel instanceof CellSelection && sel.ranges.length > 1) {
      const captured: { tr?: Transaction } = {}
      if (!mergeCells(state, (t) => (captured.tr = t)) || !captured.tr) return false
      tr = captured.tr
    }
    const table = tr.doc.nodeAt(tablePos)
    if (!table) return false
    // without the merge every selected cell is split; bottom-right first, so the
    // grid growth of one split never moves a cell still waiting its turn
    const targets: Array<{ row: number; col: number }> = []
    if (sel instanceof CellSelection && !tr.docChanged) {
      const map = TableMap.get(table)
      const seen = new Set<number>()
      for (let r = rect.top; r < rect.bottom; r++) {
        for (let c = rect.left; c < rect.right; c++) {
          const pos = map.map[r * map.width + c]
          if (seen.has(pos)) continue
          seen.add(pos)
          const cellRect = map.findCell(pos)
          targets.push({ row: cellRect.top, col: cellRect.left })
        }
      }
      targets.sort((a, b) => b.row - a.row || b.col - a.col)
    } else {
      targets.push({ row: rect.top, col: rect.left })
    }
    let current = table
    let focus: { top: number; left: number } | null = null
    for (const target of targets) {
      const result = splitGridCell(current, target, rows, cols)
      if (!result) continue
      current = result.table
      focus = result.focus
    }
    if (!focus || current === table) return false
    tr.replaceWith(tablePos, tablePos + table.nodeSize, current)
    caretIntoCell(tr, tablePos, current, focus.top, focus.left)
    dispatch?.(tr.scrollIntoView())
    return true
  }
}

/** Split Table: the selection's first row starts a new table, an empty paragraph between */
export function splitTableAtSelection(): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    const row = rect.top
    if (row <= 0) return false
    const { table } = rect
    const tablePos = rect.tableStart - 1
    const g = gridOf(table)
    const schema = table.type.schema
    const upper: GridCell[] = []
    const lower: GridCell[] = []
    for (const c of g.cells) {
      const bottom = c.top + c.rowspan
      if (bottom <= row) upper.push(c)
      else if (c.top >= row) lower.push({ ...c, top: c.top - row })
      else {
        // a merge crossing the cut: the lower part becomes an empty cell
        upper.push({ ...c, rowspan: row - c.top })
        lower.push({
          node: c.node.type.create(c.node.attrs, blankCellContent(schema, c.node)),
          top: 0,
          left: c.left,
          rowspan: bottom - row,
          colspan: c.colspan,
        })
      }
    }
    const upperTable = buildTable(table, g.rows.slice(0, row), upper)
    // formatting (style, borders, widths) carries over like Word; identity does not:
    // the lower table is a new block with no alt text of its own
    const lowerTable = buildTable(table, g.rows.slice(row), lower, {
      ...table.attrs,
      docxIndex: null,
      sdtShell: null,
      blockRevision: null,
      originalStructure: null,
      originalFormatting: null,
      tblCaption: null,
      tblDescription: null,
      tblAltEdited: false,
    })
    const gap = schema.nodes.docParagraph.create()
    const tr = state.tr.replaceWith(tablePos, tablePos + table.nodeSize, [
      upperTable,
      gap,
      lowerTable,
    ])
    caretIntoCell(tr, tablePos + upperTable.nodeSize + gap.nodeSize, lowerTable, 0, 0)
    dispatch?.(tr.scrollIntoView())
    return true
  }
}

/** rows of a multi-row selection, else the whole table */
function rowRange(rect: ReturnType<typeof selectedRect>): [number, number] {
  return rect.bottom - rect.top > 1 ? [rect.top, rect.bottom] : [0, rect.map.height]
}

function zoomOf(view: EditorView): number {
  const zoomEl = view.dom.closest('.doc-zoom') as HTMLElement | null
  return zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
}

/** Distribute Rows Evenly: the selected rows (or all) take their measured average height */
export function distributeRowsEvenly(view: EditorView): boolean {
  const { state } = view
  if (!isInTable(state)) return false
  const rect = selectedRect(state)
  const [from, to] = rowRange(rect)
  const offsets: number[] = []
  rect.table.forEach((_row, offset) => offsets.push(rect.tableStart + offset))
  const zoom = zoomOf(view)
  const heights: number[] = []
  for (let i = from; i < to; i++) {
    const dom = view.nodeDOM(offsets[i])
    if (dom instanceof HTMLElement) heights.push(dom.getBoundingClientRect().height / zoom)
  }
  const avgPx = heights.reduce((a, b) => a + b, 0) / Math.max(1, heights.length)
  if (!(avgPx > 0)) return false
  const twips = Math.round(avgPx * 15)
  let tr = state.tr
  for (let i = from; i < to; i++) {
    const row = rect.table.child(i)
    tr = tr.setNodeMarkup(offsets[i], undefined, {
      ...row.attrs,
      heightTwips: twips,
      heightRule: 'atLeast',
    })
  }
  view.dispatch(tr)
  return true
}

/** Insert N rows above/below or N columns left/right of the selection; new columns
 *  take the neighbouring column's width and the grid is refitted to the content box */
export function insertRowsOrColumns(
  count: number,
  where: 'above' | 'below' | 'left' | 'right',
  maxWidthPx: number,
): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    const tablePos = rect.tableStart - 1
    const tr = state.tr
    const n = clamp(count, 1, 100)
    if (where === 'above' || where === 'below') {
      const at = where === 'above' ? rect.top : rect.bottom
      const refIndex = where === 'above' ? rect.top : rect.bottom - 1
      const ref = rect.table.child(refIndex)
      try {
        for (let i = 0; i < n; i++) insertRowAt(tr, tablePos, at)
      } catch {
        // every column at the insertion row is covered by a vertical merge
        return false
      }
      // Word copies the adjacent row: cell formatting plus the row's own trPr flags
      copyCellFormatting(
        tr,
        tablePos,
        (r) => r.top >= at && r.top < at + n && r.bottom === r.top + 1,
        (r) => [where === 'above' ? at + n : at - 1, r.left],
      )
      const table = tr.doc.nodeAt(tablePos)!
      let rowPos = tablePos + 1
      table.forEach((row, offset, index) => {
        if (index < at || index >= at + n) return
        rowPos = tablePos + 1 + offset
        tr.setNodeMarkup(rowPos, undefined, {
          ...row.attrs,
          heightTwips: ref.attrs.heightTwips,
          heightRule: ref.attrs.heightRule,
          cantSplit: ref.attrs.cantSplit,
          cantSplitEdited: !!ref.attrs.cantSplit,
        })
      })
      dispatch?.(tr)
      return true
    }
    const at = where === 'left' ? rect.left : rect.right
    const ref = where === 'left' ? rect.left : rect.right - 1
    const widths = columnWidthsPx(rect.table, rect.map)
    const pct = pctOf(rect.table, rect.map, widths)
    const newWidths = widths ? [...widths] : null
    const newPct = [...pct]
    for (let i = 0; i < n; i++) {
      insertColumnAt(tr, tablePos, at + i)
      newWidths?.splice(at + i, 0, widths![ref])
      newPct.splice(at + i, 0, pct[ref])
    }
    copyCellFormatting(
      tr,
      tablePos,
      (r) => r.left >= at && r.left < at + n && r.right === r.left + 1,
      (r) => [r.top, where === 'left' ? at + n : at - 1],
    )
    const total = newWidths?.reduce((sum, w) => sum + w, 0) ?? 0
    // over-wide grids shrink proportionally into the content box (Word keeps the page width)
    const fitted =
      newWidths && total > maxWidthPx ? newWidths.map((w) => (w * maxWidthPx) / total) : newWidths
    reflowColumns(tr, tablePos, fitted, newPct)
    dispatch?.(tr)
    return true
  }
}

export type DeleteCellsMode = 'shiftLeft' | 'shiftUp' | 'row' | 'column'

/** shifting cells needs a plain grid: merged cells would leave holes Word refuses too */
export function canShiftCells(state: EditorState): boolean {
  if (!isInTable(state)) return false
  const rect = selectedRect(state)
  return gridOf(rect.table).cells.every((c) => c.colspan === 1 && c.rowspan === 1)
}

/** Delete Cells… dialog */
export function deleteCells(mode: DeleteCellsMode): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    if (mode === 'row') return deleteRow(state, dispatch)
    if (mode === 'column') return deleteColumn(state, dispatch)
    if (!canShiftCells(state)) return false
    const rect = selectedRect(state)
    const { table } = rect
    const tablePos = rect.tableStart - 1
    const g = gridOf(table)
    const schema = table.type.schema
    const doomed = new Set<GridCell>()
    for (const c of g.cells) {
      if (c.top >= rect.top && c.top < rect.bottom && c.left >= rect.left && c.left < rect.right)
        doomed.add(c)
    }
    if (doomed.size === 0) return false
    const width = rect.right - rect.left
    const height = rect.bottom - rect.top
    const placed: GridCell[] = []
    if (mode === 'shiftLeft') {
      // the row shortens; a borderless placeholder keeps the grid rectangular (saves as w:gridAfter)
      for (const c of g.cells) {
        if (doomed.has(c)) continue
        const inRows = c.top >= rect.top && c.top < rect.bottom
        placed.push(inRows && c.left >= rect.right ? { ...c, left: c.left - width } : c)
      }
      for (let r = rect.top; r < rect.bottom; r++) {
        for (let i = 0; i < width; i++) {
          const gone = [...doomed].find((c) => c.top === r && c.left === rect.left + i)!
          placed.push({
            node: gone.node.type.create(
              { ...gone.node.attrs, gridGap: true, fill: null, borders: null },
              blankCellContent(schema, gone.node),
            ),
            top: r,
            left: g.map.width - width + i,
            rowspan: 1,
            colspan: 1,
          })
        }
      }
    } else {
      for (const c of g.cells) {
        if (doomed.has(c)) continue
        const inCols = c.left >= rect.left && c.left < rect.right
        placed.push(inCols && c.top >= rect.bottom ? { ...c, top: c.top - height } : c)
      }
      for (let col = rect.left; col < rect.right; col++) {
        for (let i = 0; i < height; i++) {
          const last = g.cells.find((c) => c.left === col && c.top === g.map.height - 1)!
          placed.push({
            node: last.node.type.create(last.node.attrs, blankCellContent(schema, last.node)),
            top: g.map.height - height + i,
            left: col,
            rowspan: 1,
            colspan: 1,
          })
        }
      }
    }
    // rows left with placeholders only are gone (Word deletes them); an emptied
    // table becomes a paragraph, so the save never writes a w:tr without a w:tc
    const liveRows = new Set(placed.filter((c) => !c.node.attrs.gridGap).map((c) => c.top))
    if (liveRows.size === 0) {
      dispatch?.(
        state.tr.replaceWith(
          tablePos,
          tablePos + table.nodeSize,
          schema.nodes.docParagraph.create(),
        ),
      )
      return true
    }
    const keep = g.rows.map((_, i) => i).filter((i) => liveRows.has(i))
    const rowIndex = new Map(keep.map((from, to) => [from, to]))
    const kept = placed
      .filter((c) => liveRows.has(c.top))
      .map((c) => ({ ...c, top: rowIndex.get(c.top)! }))
    const next = buildTable(
      table,
      keep.map((i) => g.rows[i]),
      kept,
    )
    const tr = state.tr.replaceWith(tablePos, tablePos + table.nodeSize, next)
    caretIntoCell(
      tr,
      tablePos,
      next,
      Math.min(rect.top, keep.length - 1),
      Math.min(rect.left, g.map.width - 1),
    )
    dispatch?.(tr)
    return true
  }
}

/** every distinct cell in the selection rect: [pos of the cell node, node] */
function selectedCells(state: EditorState): Array<[number, PmNode]> {
  const rect = selectedRect(state)
  const out: Array<[number, PmNode]> = []
  const seen = new Set<number>()
  for (let row = rect.top; row < rect.bottom; row++) {
    for (let col = rect.left; col < rect.right; col++) {
      const cellPos = rect.map.map[row * rect.map.width + col]
      if (seen.has(cellPos)) continue
      seen.add(cellPos)
      const node = state.doc.nodeAt(rect.tableStart + cellPos)
      if (node && !node.attrs.gridGap) out.push([rect.tableStart + cellPos, node])
    }
  }
  return out
}

export type CellVAlign = 'top' | 'center' | 'bottom'
export type CellHAlign = 'left' | 'center' | 'right'

/** Word's 9-way cell alignment: w:vAlign on the cell plus w:jc on its paragraphs */
export function setCellAlignment(v: CellVAlign, h: CellHAlign): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    let tr = state.tr
    for (const [pos, cell] of selectedCells(state)) {
      tr = tr.setNodeMarkup(pos, undefined, { ...cell.attrs, vAlign: v === 'top' ? null : v })
      cell.forEach((child, offset) => {
        if (child.type.name !== 'docParagraph' && child.type.name !== 'docListItem') return
        tr = tr.setNodeMarkup(pos + 1 + offset, undefined, {
          ...child.attrs,
          align: alignAttrFor(h, effectiveBidi(child.attrs)),
        })
      })
    }
    dispatch?.(tr)
    return true
  }
}

export type CellTextDirection = 'lrTb' | 'tbRl' | 'btLr'

export function setCellTextDirection(dir: CellTextDirection): Command {
  return setCellProps({ textDirection: dir === 'lrTb' ? null : dir })
}

/** cell-level tcPr edits (direction / margins / noWrap / vAlign) on every selected cell */
export function setCellProps(patch: Record<string, unknown>): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    let tr = state.tr
    for (const [pos, cell] of selectedCells(state)) {
      tr = tr.setNodeMarkup(pos, undefined, { ...cell.attrs, ...patch, tcPrEdited: true })
    }
    dispatch?.(tr)
    return true
  }
}

/** trPr edits on the selected rows */
export function setRowProps(patch: Record<string, unknown>): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    let tr = state.tr
    rect.table.forEach((row, offset, index) => {
      if (index < rect.top || index >= rect.bottom) return
      tr = tr.setNodeMarkup(rect.tableStart + offset, undefined, { ...row.attrs, ...patch })
    })
    dispatch?.(tr)
    return true
  }
}

/** move the caret to the same column of another row / same row of another column (dialog prev/next) */
export function moveCaretInTable(dRow: number, dCol: number): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    const row = rect.top + dRow
    const col = rect.left + dCol
    if (row < 0 || row >= rect.map.height || col < 0 || col >= rect.map.width) return false
    const tr = state.tr
    caretIntoCell(tr, rect.tableStart - 1, rect.table, row, col)
    dispatch?.(tr)
    return true
  }
}

/** inside a table, or the whole table selected as a node (move handle / Select Table) */
export function inTableOrSelected(state: EditorState): boolean {
  const sel = state.selection
  return isInTable(state) || (sel instanceof NodeSelection && sel.node.type.name === 'docTable')
}

/** Word's Select Table for cell commands: the whole-table NodeSelection as a CellSelection over every cell */
export function tableCellsSelection(state: EditorState): CellSelection | null {
  const sel = state.selection
  if (!(sel instanceof NodeSelection) || sel.node.type.name !== 'docTable') return null
  const map = TableMap.get(sel.node)
  const start = sel.from + 1
  return CellSelection.create(state.doc, start + map.map[0], start + map.map[map.map.length - 1])
}

/** ribbon table commands on a selected table act on all of its cells (Word) */
export function enterSelectedTable(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
): boolean {
  if (isInTable(state)) return true
  const cells = tableCellsSelection(state)
  if (!cells) return false
  dispatch?.(state.tr.setSelection(cells))
  return true
}

/** the table move handle selects the whole table as a NodeSelection; cell commands need a caret inside */
export function ensureCaretInTable(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
): boolean {
  if (isInTable(state)) return true
  const sel = state.selection
  if (sel instanceof NodeSelection && sel.node.type.name === 'docTable') {
    dispatch?.(state.tr.setSelection(TextSelection.near(state.doc.resolve(sel.from + 1))))
    return true
  }
  return false
}

export type TableSelectKind = 'cell' | 'row' | 'column' | 'table'

/** Table ▸ Select: the caret's cell / row / column, or the whole table as a node */
export function selectTablePart(kind: TableSelectKind): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false
    const rect = selectedRect(state)
    if (kind === 'table') {
      dispatch?.(state.tr.setSelection(NodeSelection.create(state.doc, rect.tableStart - 1)))
      return true
    }
    const $cell = state.doc.resolve(
      rect.tableStart + rect.map.map[rect.top * rect.map.width + rect.left],
    )
    const selection =
      kind === 'row'
        ? CellSelection.rowSelection($cell)
        : kind === 'column'
          ? CellSelection.colSelection($cell)
          : new CellSelection($cell)
    dispatch?.(state.tr.setSelection(selection))
    return true
  }
}

/** table dialogs and menu commands change the document: refused with a hint on a read-only editor */
export function editableOrToast(editor: { isEditable: boolean }): boolean {
  if (editor.isEditable) return true
  showToast(t('ribbonReadOnlyEdit'), 'error')
  return false
}
