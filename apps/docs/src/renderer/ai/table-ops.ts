import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import { EditorState } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import {
  CellSelection,
  TableMap,
  addColSpan,
  removeColumn,
  removeRow,
  rowIsHeader,
  splitCell,
  tableNodeTypes,
  type Rect,
  type TableRect,
} from '@tiptap/pm/tables'
import type { CellBorder, CellBorders, StyleInfo, TableLook } from '@genoffice/docx-engine'
import type { Op, OpDef, OpResult, RunEnv, Target, TopBlock } from './ops'

/**
 * Structural table ops: rows, columns, merges, cell formatting and the table
 * style. Row/column indexes are 0-based grid coordinates (a merged cell counts
 * once per grid slot it covers). Every op targets exactly one native table
 * block; the save path regenerates the table from the model, so gridSpan /
 * vMerge / tcPr come out consistent by construction.
 */

export interface TableOpHelpers {
  validateShape(op: Op, def: OpDef, where: string): string | null
  matchTarget(doc: PmNode, target: Target, sel: RunEnv['sel']): TopBlock[]
  markChanged(tr: Transaction, pos: number, ctx: RunEnv['ctx']): void
  normalizeHex(value: unknown, where: string, field: string): string | null
  checkHex(value: unknown, where: string, field: string): string | null
}

interface CellRange {
  row: number
  col: number
  rowEnd?: number
  colEnd?: number
}

interface BorderInput {
  style?: string
  /** points */
  width?: number
  color?: string
}

const V_ALIGNS = ['top', 'center', 'bottom'] as const
const BORDER_SIDES = ['top', 'left', 'bottom', 'right'] as const
const LOOK_KEYS: (keyof TableLook)[] = [
  'firstRow',
  'lastRow',
  'firstColumn',
  'lastColumn',
  'bandedRows',
  'bandedColumns',
]
const TWIPS_PER: Record<string, number> = { twip: 1, pt: 20, px: 15, in: 1440, cm: 567, mm: 56.7 }
const LENGTH = /^\s*(\d+(?:\.\d+)?)\s*(twip|pt|px|in|cm|mm)\s*$/i
/** Word's default body width (Letter, 1in margins); only used for tables without any width */
const DEFAULT_BODY_TWIPS = 9360

/** number = twips; "2.5cm" / "1in" / "72pt" / "96px" / "10mm" */
export function parseTableLength(value: unknown): number | undefined {
  if (typeof value === 'number') return value > 0 ? value : undefined
  if (typeof value !== 'string') return undefined
  const m = LENGTH.exec(value)
  if (!m) return undefined
  const twips = Number(m[1]) * TWIPS_PER[m[2]!.toLowerCase()]!
  return twips > 0 ? twips : undefined
}

interface TableHit {
  node: PmNode
  pos: number
  map: TableMap
}

function resolveTable(op: Op, env: RunEnv, helpers: TableOpHelpers): TableHit | null {
  const blocks = helpers.matchTarget(env.tr.doc, op.target!, env.sel)
  if (blocks.length === 0) return null
  const tables = blocks.filter((b) => b.node.type.name === 'docTable')
  if (tables.length === 0) {
    const kinds = blocks.map((b) =>
      b.node.type.name === 'docProtected'
        ? `read-only ${b.node.attrs.blockType}`
        : b.node.type.name,
    )
    throw new Error(
      `${op.op}: the target matches no native table (matched: ${kinds.join(', ')}); use a target with nodeType "table"`,
    )
  }
  if (tables.length > 1) {
    throw new Error(
      `${op.op}: the target matches ${tables.length} tables (blocks ${tables.map((b) => b.index).join(', ')}); narrow it to one with blockIndexes`,
    )
  }
  const { node, pos } = tables[0]
  return { node, pos, map: TableMap.get(node) }
}

const isInt = (v: unknown): v is number => Number.isInteger(v) && Number(v) >= 0

function checkIndex(op: Op, where: string, key: string): string | null {
  return isInt(op[key]) ? null : `${where}: ${key} must be a non-negative integer (0-based)`
}

function checkCount(op: Op, where: string): string | null {
  if (op.count === undefined) return null
  return Number.isInteger(op.count) && Number(op.count) >= 1 && Number(op.count) <= 100
    ? null
    : `${where}: count must be an integer between 1 and 100`
}

function checkPosition(op: Op, where: string): string | null {
  if (op.position === undefined || op.position === 'before' || op.position === 'after') return null
  return `${where}: position must be "before" or "after"`
}

function checkRange(value: unknown, where: string, key: string): string | null {
  if (!value || typeof value !== 'object')
    return `${where}: ${key} must be { row, col, rowEnd?, colEnd? }`
  const r = value as Record<string, unknown>
  for (const k of ['row', 'col']) {
    if (!isInt(r[k])) return `${where}: ${key}.${k} must be a non-negative integer (0-based)`
  }
  for (const k of ['rowEnd', 'colEnd']) {
    if (r[k] !== undefined && !isInt(r[k])) {
      return `${where}: ${key}.${k} must be a non-negative integer (0-based)`
    }
  }
  if (r.rowEnd !== undefined && Number(r.rowEnd) < Number(r.row)) {
    return `${where}: ${key}.rowEnd must be >= row`
  }
  if (r.colEnd !== undefined && Number(r.colEnd) < Number(r.col)) {
    return `${where}: ${key}.colEnd must be >= col`
  }
  return null
}

function rectOf(range: CellRange | undefined, map: TableMap, op: string): Rect {
  if (!range) return { left: 0, top: 0, right: map.width, bottom: map.height }
  const top = range.row
  const left = range.col
  const bottom = (range.rowEnd ?? range.row) + 1
  const right = (range.colEnd ?? range.col) + 1
  if (bottom > map.height) {
    throw new Error(`${op}: row ${bottom - 1} out of range (0-${map.height - 1})`)
  }
  if (right > map.width) {
    throw new Error(`${op}: col ${right - 1} out of range (0-${map.width - 1})`)
  }
  return { left, top, right, bottom }
}

/** cell positions (relative to the table start) covering a rect, each once */
/** First row in [rowFrom, rowTo) that keeps no cell of its own once cols [colFrom, colTo) go away, or -1. */
function rowLeftEmpty(
  map: TableMap,
  rowFrom: number,
  rowTo: number,
  colFrom: number,
  colTo: number,
): number {
  for (let row = rowFrom; row < rowTo; row++) {
    let survives = false
    for (let col = 0; col < map.width && !survives; col++) {
      if (col >= colFrom && col < colTo) continue
      survives = map.findCell(map.map[row * map.width + col]!).top === row
    }
    if (!survives) return row
  }
  return -1
}

function cellsInRect(map: TableMap, rect: Rect): number[] {
  const seen = new Set<number>()
  for (let row = rect.top; row < rect.bottom; row++) {
    for (let col = rect.left; col < rect.right; col++) {
      seen.add(map.map[row * map.width + col]!)
    }
  }
  return [...seen]
}

/** px per grid column when the table carries absolute widths */
function columnWidthsPx(table: PmNode, map: TableMap): number[] | null {
  const widths: number[] = []
  const first = table.firstChild
  if (!first) return null
  for (let i = 0; i < first.childCount; i++) {
    const cw = first.child(i).attrs.colwidth as number[] | null
    if (!cw?.length) return null
    widths.push(...cw)
  }
  if (widths.length !== map.width || widths.some((w) => !(w > 0))) return null
  return widths
}

/** widths to fall back to when a table has no px grid: its pct share of the page body */
function derivedWidthsPx(table: PmNode, map: TableMap): number[] {
  const pct = table.attrs.colWidthsPct as number[] | null
  const bodyPx = DEFAULT_BODY_TWIPS / 15
  if (pct?.length === map.width) {
    const total = pct.reduce((s, p) => s + Math.max(0, p), 0) || 100
    return pct.map((p) => Math.max(1, Math.round((Math.max(0, p) / total) * bodyPx)))
  }
  return Array.from({ length: map.width }, () => Math.round(bodyPx / map.width))
}

/**
 * Re-derive every cell's colwidth and the table's width attrs from one widths
 * array (px per grid column). `widths === null` keeps the table percentage
 * based: cells lose their px widths and colWidthsPct is renormalized.
 */
function reflowColumns(
  tr: Transaction,
  tablePos: number,
  widths: number[] | null,
  pct: number[],
): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const seen = new Set<number>()
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i]!
    if (seen.has(pos)) continue
    seen.add(pos)
    const cell = table.nodeAt(pos)!
    const rect = map.findCell(pos)
    const colwidth = widths ? widths.slice(rect.left, rect.right) : null
    if (JSON.stringify(colwidth) === JSON.stringify(cell.attrs.colwidth)) continue
    tr.setNodeMarkup(start + pos, undefined, { ...cell.attrs, colwidth })
  }
  const total = pct.reduce((s, p) => s + p, 0) || 1
  tr.setNodeMarkup(tablePos, undefined, {
    ...table.attrs,
    colWidthsPct: pct.map((p) => (p / total) * 100),
    widthPx:
      widths && table.attrs.widthPx != null
        ? widths.reduce((s, w) => s + w, 0)
        : table.attrs.widthPx,
  })
}

function pctOf(table: PmNode, map: TableMap, widths: number[] | null): number[] {
  if (widths) {
    const total = widths.reduce((s, w) => s + w, 0) || 1
    return widths.map((w) => (w / total) * 100)
  }
  const pct = table.attrs.colWidthsPct as number[] | null
  return pct?.length === map.width
    ? [...pct]
    : Array.from({ length: map.width }, () => 100 / map.width)
}

/** prosemirror-tables' addColumn maps through the whole transaction; this one only through its own steps */
function insertColumnAt(tr: Transaction, tablePos: number, col: number): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const mapStart = tr.mapping.maps.length
  const refColumn = col > 0 ? -1 : 0
  const cellType = tableNodeTypes(table.type.schema).cell
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + col
    if (col > 0 && col < map.width && map.map[index - 1] === map.map[index]) {
      const pos = map.map[index]!
      const cell = table.nodeAt(pos)!
      tr.setNodeMarkup(
        tr.mapping.slice(mapStart).map(start + pos),
        undefined,
        addColSpan(cell.attrs as Parameters<typeof addColSpan>[0], col - map.colCount(pos)),
      )
      row += Number(cell.attrs.rowspan) - 1
    } else {
      const refPos = map.map[index + refColumn]
      const type = refPos === undefined ? cellType : table.nodeAt(refPos)!.type
      const pos = map.positionAt(row, col, table)
      tr.insert(tr.mapping.slice(mapStart).map(start + pos), type.createAndFill()!)
    }
  }
}

const COPIED_CELL_ATTRS = [
  'fill',
  'borders',
  'vAlign',
  'cellMar',
  'textDirection',
  'align',
  'bold',
  'color',
] as const

/** new rows/columns take the visual attrs of the neighbouring cell, like Word */
function copyCellFormatting(
  tr: Transaction,
  tablePos: number,
  isNew: (rect: Rect) => boolean,
  refOf: (rect: Rect) => [number, number],
): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  const seen = new Set<number>()
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i]!
    if (seen.has(pos)) continue
    seen.add(pos)
    const rect = map.findCell(pos)
    if (!isNew(rect)) continue
    const [refRow, refCol] = refOf(rect)
    if (refRow < 0 || refRow >= map.height || refCol < 0 || refCol >= map.width) continue
    const ref = table.nodeAt(map.map[refRow * map.width + refCol]!)!
    const cell = table.nodeAt(pos)!
    const patch: Record<string, unknown> = {}
    for (const k of COPIED_CELL_ATTRS) patch[k] = ref.attrs[k]
    tr.setNodeMarkup(start + pos, undefined, { ...cell.attrs, ...patch })
  }
}

/** prosemirror-tables' addRow, but positions are read from the current doc, not mapped through earlier batch steps */
function insertRowAt(tr: Transaction, tablePos: number, row: number): void {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  const start = tablePos + 1
  let rowPos = start
  for (let i = 0; i < row; i++) rowPos += table.child(i).nodeSize
  const cells: PmNode[] = []
  let refRow: number | null = row > 0 ? -1 : 0
  if (rowIsHeader(map, table, row + refRow)) refRow = row === 0 || row === map.height ? null : 0
  const types = tableNodeTypes(table.type.schema)
  for (let col = 0; col < map.width; col++) {
    const index = map.width * row + col
    if (row > 0 && row < map.height && map.map[index] === map.map[index - map.width]) {
      const pos = map.map[index]!
      const attrs = table.nodeAt(pos)!.attrs
      tr.setNodeMarkup(start + pos, undefined, { ...attrs, rowspan: Number(attrs.rowspan) + 1 })
      col += Number(attrs.colspan) - 1
    } else {
      const ref = refRow === null ? undefined : map.map[index + refRow * map.width]
      const type = ref === undefined ? types.cell : table.nodeAt(ref)!.type
      cells.push(type.createAndFill()!)
    }
  }
  if (cells.length === 0) {
    throw new Error(
      `insertTableRow: every column at row ${row} is covered by vertically merged cells, so no new cell can be placed there; insert outside the merged rows or splitTableCell first`,
    )
  }
  tr.insert(rowPos, types.row.create(null, cells))
}

/** prosemirror-tables' row/column helpers take the whole table as their rect */
function tableRect(tr: Transaction, tablePos: number): TableRect {
  const table = tr.doc.nodeAt(tablePos)!
  const map = TableMap.get(table)
  return {
    map,
    table,
    tableStart: tablePos + 1,
    left: 0,
    top: 0,
    right: map.width,
    bottom: map.height,
  }
}

function tableResult(op: string, changed: number): OpResult {
  return { op, matched: 1, changed, skippedProtected: 0 }
}

const noMatch = (op: string): OpResult => ({ op, matched: 0, changed: 0, skippedProtected: 0 })

export function tableOpDefs(h: TableOpHelpers): OpDef[] {
  const tableTarget = (op: Op, env: RunEnv) => resolveTable(op, env, h)

  function runInsertRow(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos } = hit
    const row = Number(op.row)
    if (row >= map.height)
      throw new Error(`${op.op}: row ${row} out of range (0-${map.height - 1})`)
    const count = Number(op.count ?? 1)
    const at = op.position === 'before' ? row : row + 1
    const widths = columnWidthsPx(hit.node, map)
    const pct = pctOf(hit.node, map, widths)
    for (let i = 0; i < count; i++) {
      insertRowAt(env.tr, pos, at)
    }
    copyCellFormatting(
      env.tr,
      pos,
      (rect) => rect.top >= at && rect.top < at + count && rect.bottom === rect.top + 1,
      (rect) => [op.position === 'before' ? at + count : at - 1, rect.left],
    )
    reflowColumns(env.tr, pos, widths, pct)
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, count)
  }

  function runDeleteRow(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos } = hit
    const row = Number(op.row)
    const count = Number(op.count ?? 1)
    if (row + count > map.height) {
      throw new Error(`${op.op}: rows ${row}-${row + count - 1} out of range (0-${map.height - 1})`)
    }
    if (count >= map.height) {
      throw new Error(
        `${op.op}: cannot delete every row of the table; use deleteBlocks to remove it`,
      )
    }
    for (let i = 0; i < count; i++) {
      removeRow(env.tr, tableRect(env.tr, pos), row)
    }
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, count)
  }

  function runInsertColumn(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos } = hit
    const col = Number(op.col)
    if (col >= map.width) throw new Error(`${op.op}: col ${col} out of range (0-${map.width - 1})`)
    const count = Number(op.count ?? 1)
    const at = op.position === 'before' ? col : col + 1
    const widths = columnWidthsPx(hit.node, map)
    const pct = pctOf(hit.node, map, widths)
    const refIdx = op.position === 'before' ? col : col
    const newWidths = widths ? [...widths] : null
    const newPct = [...pct]
    for (let i = 0; i < count; i++) {
      insertColumnAt(env.tr, pos, at + i)
      newWidths?.splice(at + i, 0, widths![refIdx]!)
      newPct.splice(at + i, 0, pct[refIdx]!)
    }
    copyCellFormatting(
      env.tr,
      pos,
      (rect) => rect.left >= at && rect.left < at + count && rect.right === rect.left + 1,
      (rect) => [rect.top, op.position === 'before' ? at + count : at - 1],
    )
    reflowColumns(env.tr, pos, newWidths, newPct)
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, count)
  }

  function runDeleteColumn(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos } = hit
    const col = Number(op.col)
    const count = Number(op.count ?? 1)
    if (col + count > map.width) {
      throw new Error(`${op.op}: cols ${col}-${col + count - 1} out of range (0-${map.width - 1})`)
    }
    if (count >= map.width) {
      throw new Error(
        `${op.op}: cannot delete every column of the table; use deleteBlocks to remove it`,
      )
    }
    const emptied = rowLeftEmpty(map, 0, map.height, col, col + count)
    if (emptied >= 0) {
      throw new Error(
        `${op.op}: deleting cols ${col}-${col + count - 1} would leave row ${emptied} without cells (its other columns are covered by a rowspan from above); split that merge or use deleteTableRow`,
      )
    }
    const widths = columnWidthsPx(hit.node, map)
    const pct = pctOf(hit.node, map, widths)
    for (let i = 0; i < count; i++) {
      removeColumn(env.tr, tableRect(env.tr, pos), col)
    }
    widths?.splice(col, count)
    pct.splice(col, count)
    reflowColumns(env.tr, pos, widths, pct)
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, count)
  }

  function runMergeCells(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos, node: table } = hit
    const rect = rectOf(op.range as CellRange, map, op.op)
    if (rect.right - rect.left === 1 && rect.bottom - rect.top === 1) {
      throw new Error(`${op.op}: range covers a single cell; give rowEnd and/or colEnd`)
    }
    const cells = cellsInRect(map, rect)
    const emptied = rowLeftEmpty(map, rect.top + 1, rect.bottom, rect.left, rect.right)
    if (emptied >= 0) {
      throw new Error(
        `${op.op}: merging rows ${rect.top}-${rect.bottom - 1} across the whole width would leave row ${emptied} without cells; use deleteTableRow instead`,
      )
    }
    for (const cellPos of cells) {
      const r = map.findCell(cellPos)
      if (
        r.left < rect.left ||
        r.right > rect.right ||
        r.top < rect.top ||
        r.bottom > rect.bottom
      ) {
        throw new Error(
          `${op.op}: the cell at row ${r.top} col ${r.left} spans rows ${r.top}-${r.bottom - 1}, cols ${r.left}-${r.right - 1} and crosses the range; include it whole or choose another range`,
        )
      }
    }
    const start = pos + 1
    const mapStart = env.tr.mapping.maps.length
    const [firstPos, ...rest] = cells.sort((a, b) => a - b)
    const first = table.nodeAt(firstPos!)!
    let content = Fragment.empty
    const firstEmpty = isEmptyCell(first)
    for (const cellPos of rest) {
      const cell = table.nodeAt(cellPos)!
      if (!isEmptyCell(cell)) content = content.append(cell.content)
      const from = env.tr.mapping.slice(mapStart).map(start + cellPos)
      env.tr.delete(from, from + cell.nodeSize)
    }
    const mergedPos = env.tr.mapping.slice(mapStart).map(start + firstPos!)
    const widths = columnWidthsPx(table, map)
    env.tr.setNodeMarkup(mergedPos, undefined, {
      ...first.attrs,
      colspan: rect.right - rect.left,
      rowspan: rect.bottom - rect.top,
      colwidth: widths ? widths.slice(rect.left, rect.right) : null,
    })
    if (content.size > 0) {
      const merged = env.tr.doc.nodeAt(mergedPos)!
      const contentStart = mergedPos + 1
      const contentEnd = contentStart + merged.content.size
      env.tr.replaceWith(firstEmpty ? contentStart : contentEnd, contentEnd, content)
    }
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, cells.length)
  }

  function runSplitCell(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos, node: table } = hit
    const rect = rectOf(op.cell as CellRange, map, op.op)
    const cellPos = map.map[rect.top * map.width + rect.left]!
    const cell = table.nodeAt(cellPos)!
    if (Number(cell.attrs.colspan) === 1 && Number(cell.attrs.rowspan) === 1) {
      throw new Error(`${op.op}: the cell at row ${rect.top} col ${rect.left} is not merged`)
    }
    const abs = pos + 1 + cellPos
    const state = EditorState.create({
      doc: env.tr.doc,
      selection: CellSelection.create(env.tr.doc, abs),
    })
    let applied = false
    splitCell(state, (t) => {
      for (const step of t.steps) env.tr.step(step)
      applied = true
    })
    if (!applied)
      throw new Error(`${op.op}: the cell at row ${rect.top} col ${rect.left} cannot be split`)
    const widths = columnWidthsPx(table, map)
    reflowColumns(env.tr, pos, widths, pctOf(table, map, widths))
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, Number(cell.attrs.colspan) * Number(cell.attrs.rowspan))
  }

  function runSetCellFormat(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { map, pos, node: table } = hit
    const rect = rectOf(op.range as CellRange | undefined, map, op.op)
    const cells = cellsInRect(map, rect)
    const start = pos + 1
    const fill = op.fill === undefined ? undefined : h.normalizeHex(op.fill, op.op, 'fill')
    for (const cellPos of cells) {
      const cell = table.nodeAt(cellPos)!
      const patch: Record<string, unknown> = {}
      if (fill !== undefined) patch.fill = fill
      if (op.vAlign !== undefined) patch.vAlign = op.vAlign === 'top' ? null : op.vAlign
      if (op.borders !== undefined) {
        patch.borders = mergeBorders(cell.attrs.borders as CellBorders | null, op.borders, h)
      }
      if (Object.keys(patch).length > 0) {
        env.tr.setNodeMarkup(start + cellPos, undefined, { ...cell.attrs, ...patch })
      }
    }
    if (op.width !== undefined) {
      const px = Math.max(1, Math.round(parseTableLength(op.width)! / 15))
      const widths = columnWidthsPx(table, map) ?? derivedWidthsPx(table, map)
      for (let c = rect.left; c < rect.right; c++) widths[c] = px
      reflowColumns(env.tr, pos, widths, pctOf(table, map, widths))
    }
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, cells.length)
  }

  function runSetTableStyle(op: Op, env: RunEnv): OpResult {
    const hit = tableTarget(op, env)
    if (!hit) return noMatch(op.op)
    const { pos, node: table } = hit
    const patch: Record<string, unknown> = {}
    if (op.styleId !== undefined) {
      if (op.styleId === null) {
        patch.tblStyleId = null
      } else {
        const styles = tableStyles(env)
        const id = String(op.styleId)
        const match =
          styles.find((s) => s.styleId === id) ??
          styles.find((s) => s.name.toLowerCase() === id.toLowerCase())
        if (!match) {
          throw new Error(
            `${op.op}: no table style "${id}" in this document. Available: [${styles.map((s) => s.styleId).join(', ')}]`,
          )
        }
        patch.tblStyleId = match.styleId
      }
    }
    if (op.look !== undefined) {
      const current = (table.attrs.tblLook as TableLook | null) ?? {
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        bandedRows: true,
        bandedColumns: false,
      }
      patch.tblLook = { ...current, ...(op.look as Partial<TableLook>) }
      patch.tblLookEdited = true
    }
    env.tr.setNodeMarkup(pos, undefined, { ...table.attrs, ...patch })
    h.markChanged(env.tr, pos, env.ctx)
    return tableResult(op.op, 1)
  }

  const rowSig = (name: string, extra: string) =>
    `{ op: "${name}", target, row, ${extra} }  // target must match one native table (nodeType "table" and/or blockIndexes); row is 0-based`
  const colSig = (name: string, extra: string) =>
    `{ op: "${name}", target, col, ${extra} }  // target must match one native table; col is a 0-based grid column`

  return [
    {
      name: 'insertTableRow',
      signature:
        rowSig('insertTableRow', 'position?: "before"|"after", count?') +
        "; new rows copy the neighbouring row's cell shading/borders (default: after)",
      keys: ['row', 'position', 'count'],
      target: 'required',
      validate(op, where) {
        return (
          h.validateShape(op, this, where) ??
          checkIndex(op, where, 'row') ??
          checkPosition(op, where) ??
          checkCount(op, where)
        )
      },
      apply: runInsertRow,
    },
    {
      name: 'deleteTableRow',
      signature: rowSig('deleteTableRow', 'count?') + '; cells merged across the row shrink',
      keys: ['row', 'count'],
      target: 'required',
      validate(op, where) {
        return (
          h.validateShape(op, this, where) ?? checkIndex(op, where, 'row') ?? checkCount(op, where)
        )
      },
      apply: runDeleteRow,
    },
    {
      name: 'insertTableColumn',
      signature:
        colSig('insertTableColumn', 'position?: "before"|"after", count?') +
        "; the new column takes the neighbouring column's width and cell formatting (default: after)",
      keys: ['col', 'position', 'count'],
      target: 'required',
      validate(op, where) {
        return (
          h.validateShape(op, this, where) ??
          checkIndex(op, where, 'col') ??
          checkPosition(op, where) ??
          checkCount(op, where)
        )
      },
      apply: runInsertColumn,
    },
    {
      name: 'deleteTableColumn',
      signature: colSig('deleteTableColumn', 'count?') + '; cells merged across the column shrink',
      keys: ['col', 'count'],
      target: 'required',
      validate(op, where) {
        return (
          h.validateShape(op, this, where) ?? checkIndex(op, where, 'col') ?? checkCount(op, where)
        )
      },
      apply: runDeleteColumn,
    },
    {
      name: 'mergeTableCells',
      signature:
        '{ op: "mergeTableCells", target, range: { row, col, rowEnd?, colEnd? } }  // merge the rectangle (0-based grid coordinates, ends inclusive) into one cell; refused when an existing merged cell crosses the rectangle edge; the texts are kept as separate paragraphs',
      keys: ['range'],
      target: 'required',
      validate(op, where) {
        return h.validateShape(op, this, where) ?? checkRange(op.range, where, 'range')
      },
      apply: runMergeCells,
    },
    {
      name: 'splitTableCell',
      signature:
        '{ op: "splitTableCell", target, cell: { row, col } }  // undo a merge: the cell becomes one cell per grid slot it covered; the text stays in the top-left one',
      keys: ['cell'],
      target: 'required',
      validate(op, where) {
        return h.validateShape(op, this, where) ?? checkRange(op.cell, where, 'cell')
      },
      apply: runSplitCell,
    },
    {
      name: 'setTableCellFormat',
      signature:
        '{ op: "setTableCellFormat", target, range?: { row, col, rowEnd?, colEnd? }, fill?: "#RRGGBB"|null, vAlign?: "top"|"center"|"bottom"|null, width?: twips|"2.5cm"|"1in"|"72pt"|"96px", borders?: { all?, top?, left?, bottom?, right?: { style?: "single"|"dashed"|"double"|"none", width?: pt, color?: "#RRGGBB" }|null }|null }  // cells in the range (omitted = whole table); width sets the grid columns the range covers; borders: null = no borders on the cell, a side null = that side off',
      keys: ['range', 'fill', 'vAlign', 'width', 'borders'],
      target: 'required',
      validate(op, where) {
        const shape = h.validateShape(op, this, where)
        if (shape) return shape
        if (op.range !== undefined) {
          const r = checkRange(op.range, where, 'range')
          if (r) return r
        }
        if (op.fill !== undefined) {
          const hex = h.checkHex(op.fill, where, 'fill')
          if (hex) return hex
        }
        if (
          op.vAlign !== undefined &&
          op.vAlign !== null &&
          !V_ALIGNS.includes(op.vAlign as never)
        ) {
          return `${where}: vAlign must be top / center / bottom (or null)`
        }
        if (op.width !== undefined && parseTableLength(op.width) === undefined) {
          return `${where}: width must be a positive number of twips or a length like "2.5cm", "1in", "72pt", "96px"`
        }
        if (op.borders !== undefined) {
          const b = checkBorders(op.borders, where, h)
          if (b) return b
        }
        if (['fill', 'vAlign', 'width', 'borders'].every((k) => op[k] === undefined)) {
          return `${where}: give at least one of fill, vAlign, width, borders`
        }
        return null
      },
      apply: runSetCellFormat,
    },
    {
      name: 'setTableStyle',
      signature:
        '{ op: "setTableStyle", target, styleId?: string|null, look?: { firstRow?, lastRow?, firstColumn?, lastColumn?, bandedRows?, bandedColumns?: boolean } }  // styleId must be a table style defined in this document (docs read lists them as tableStyles); null removes the style; look toggles Word\'s Table Style Options',
      keys: ['styleId', 'look'],
      target: 'required',
      validate(op, where) {
        const shape = h.validateShape(op, this, where)
        if (shape) return shape
        if (op.styleId !== undefined && op.styleId !== null && typeof op.styleId !== 'string') {
          return `${where}: styleId must be a string (or null)`
        }
        if (op.look !== undefined) {
          if (!op.look || typeof op.look !== 'object')
            return `${where}: look must be an object of booleans`
          for (const [k, v] of Object.entries(op.look as Record<string, unknown>)) {
            if (!LOOK_KEYS.includes(k as keyof TableLook)) {
              return `${where}: unknown look key "${k}"; allowed: ${LOOK_KEYS.join(', ')}`
            }
            if (typeof v !== 'boolean') return `${where}: look.${k} must be a boolean`
          }
        }
        if (op.styleId === undefined && op.look === undefined) {
          return `${where}: give styleId and/or look`
        }
        return null
      },
      apply: runSetTableStyle,
    },
  ]
}

function isEmptyCell(cell: PmNode): boolean {
  const c = cell.content
  return c.childCount === 1 && c.firstChild!.isTextblock && c.firstChild!.content.size === 0
}

function tableStyles(env: RunEnv): StyleInfo[] {
  const styles = (env.editor.storage as { listNumbering?: { styles?: Map<string, StyleInfo> } })
    .listNumbering?.styles
  return [...(styles?.values() ?? [])].filter((s) => s.type === 'table')
}

function checkBorders(value: unknown, where: string, h: TableOpHelpers): string | null {
  if (value === null) return null
  if (!value || typeof value !== 'object') return `${where}: borders must be an object (or null)`
  for (const [side, b] of Object.entries(value as Record<string, unknown>)) {
    if (side !== 'all' && !BORDER_SIDES.includes(side as never)) {
      return `${where}: unknown borders side "${side}"; allowed: all, ${BORDER_SIDES.join(', ')}`
    }
    if (b === null) continue
    if (!b || typeof b !== 'object') return `${where}: borders.${side} must be an object (or null)`
    const bi = b as BorderInput
    if (bi.style !== undefined && (typeof bi.style !== 'string' || !/^[a-zA-Z]+$/.test(bi.style))) {
      return `${where}: borders.${side}.style must be a Word border name (single, dashed, double, none, …)`
    }
    if (
      bi.width !== undefined &&
      !(typeof bi.width === 'number' && bi.width > 0 && bi.width <= 96)
    ) {
      return `${where}: borders.${side}.width must be points between 0 and 96`
    }
    if (bi.color !== undefined) {
      const hex = h.checkHex(bi.color, where, `borders.${side}.color`)
      if (hex) return hex
    }
  }
  return null
}

function mergeBorders(
  current: CellBorders | null,
  input: unknown,
  h: TableOpHelpers,
): CellBorders | null {
  const none: CellBorder = { style: 'none' }
  if (input === null) return { top: none, left: none, bottom: none, right: none }
  const spec = input as Record<string, BorderInput | null | undefined>
  const out: CellBorders = { ...(current ?? {}) }
  const toBorder = (side: string, b: BorderInput | null): CellBorder => {
    if (b === null) return none
    const base = (current?.[side as keyof CellBorders] as CellBorder | undefined) ?? {
      style: 'single',
    }
    return {
      style: b.style ?? (base.style === 'none' ? 'single' : base.style),
      ...(b.width !== undefined
        ? { szEighths: Math.round(b.width * 8) }
        : base.szEighths !== undefined
          ? { szEighths: base.szEighths }
          : { szEighths: 4 }),
      ...(b.color !== undefined
        ? { color: h.normalizeHex(b.color, 'borders', 'color') ?? 'auto' }
        : base.color !== undefined
          ? { color: base.color }
          : { color: 'auto' }),
    }
  }
  for (const side of BORDER_SIDES) {
    const own = spec[side]
    const all = spec.all
    if (own !== undefined) out[side] = toBorder(side, own)
    else if (all !== undefined) out[side] = toBorder(side, all)
  }
  return out
}

export interface TableSummary {
  rows: number
  cols: number
  merged?: Array<{ row: number; col: number; rowSpan: number; colSpan: number }>
  styleId?: string
}

/** dimensions and merged cells, for agents that address grid coordinates */
export function describeTable(table: PmNode): TableSummary {
  const map = TableMap.get(table)
  const merged: NonNullable<TableSummary['merged']> = []
  const seen = new Set<number>()
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i]!
    if (seen.has(pos)) continue
    seen.add(pos)
    const cell = table.nodeAt(pos)!
    const colSpan = Number(cell.attrs.colspan) || 1
    const rowSpan = Number(cell.attrs.rowspan) || 1
    if (colSpan > 1 || rowSpan > 1) {
      const rect = map.findCell(pos)
      merged.push({ row: rect.top, col: rect.left, rowSpan, colSpan })
    }
  }
  const styleId = table.attrs.tblStyleId as string | null
  return {
    rows: map.height,
    cols: map.width,
    ...(merged.length ? { merged } : {}),
    ...(styleId ? { styleId } : {}),
  }
}
