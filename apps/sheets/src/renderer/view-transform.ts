/// Screen ↔ file coordinate mapping for streamed sheets with journaled
/// structural operations. "File" space is the original xlsx the sidecar
/// reads; "screen" space is Univer's post-operation model. Viewport requests
/// translate screen → file, streamed results translate file → screen; rows
/// and columns inserted this session have no file backing (`null`).

import type { StructuralOp } from '../gateway/xlsx-structure'
import type { WorkbookRangeResult } from '../shared/desktop-api'

export type Axis = 'row' | 'column'

interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

type RowColumnOp = Extract<StructuralOp, { index: number }>

function axisOf(op: RowColumnOp): Axis {
  return op.kind === 'insert-cols' || op.kind === 'remove-cols' ? 'column' : 'row'
}

/// The two adjacent pre-move blocks a move swaps: `first` then `second`,
/// with `second` immediately following `first`.
function swapBlocks(op: Extract<RowColumnOp, { before: number }>): {
  first: { start: number; end: number }
  second: { start: number; end: number }
} {
  const first =
    op.before > op.index
      ? { start: op.index, end: op.index + op.count - 1 }
      : { start: op.before, end: op.index - 1 }
  const second =
    op.before > op.index
      ? { start: op.index + op.count, end: op.before - 1 }
      : { start: op.index, end: op.index + op.count - 1 }
  return { first, second }
}

/// The two blocks a move swaps, in the coordinate space the map operates on:
/// pre-move blocks for the forward map, their post-move images (split at
/// `first.start + secondLength`) for the inverse.
function swapImageBlocks(
  op: Extract<RowColumnOp, { before: number }>,
  forward: boolean,
): [{ start: number; end: number }, { start: number; end: number }] {
  const { first, second } = swapBlocks(op)
  const secondLength = second.end - second.start + 1
  return forward
    ? [first, second]
    : [
        { start: first.start, end: first.start + secondLength - 1 },
        { start: first.start + secondLength, end: second.end },
      ]
}

/// A move as the two adjacent blocks that swap places; `forward` maps
/// pre-move → post-move, its inverse swaps the (equal-length) images back.
function swapMap(
  position: number,
  op: Extract<RowColumnOp, { before: number }>,
  forward: boolean,
): number {
  const [a, b] = swapImageBlocks(op, forward)
  if (position >= a.start && position <= a.end) {
    return position + (b.end - b.start + 1)
  }
  if (position >= b.start && position <= b.end) {
    return position - (a.end - a.start + 1)
  }
  return position
}

function isInsert(op: RowColumnOp): boolean {
  return op.kind === 'insert-rows' || op.kind === 'insert-cols'
}

function rowColumnOps(ops: readonly StructuralOp[], axis: Axis): RowColumnOp[] {
  return ops.filter((op): op is RowColumnOp => 'index' in op && axisOf(op) === axis)
}

/// Where a file line sits on screen after all operations, or null when a
/// removal deleted it.
export function fileToScreen(
  ops: readonly StructuralOp[],
  axis: Axis,
  index: number,
): number | null {
  let position = index
  for (const op of rowColumnOps(ops, axis)) {
    if ('before' in op) {
      position = swapMap(position, op, true)
    } else if (isInsert(op)) {
      if (position >= op.index) position += op.count
    } else {
      if (position >= op.index && position < op.index + op.count) return null
      if (position >= op.index + op.count) position -= op.count
    }
  }
  return position
}

/// Which file line backs a screen position, or null when the line was
/// inserted this session (journal-owned, nothing to stream).
export function screenToFile(
  ops: readonly StructuralOp[],
  axis: Axis,
  index: number,
): number | null {
  let position = index
  const relevant = rowColumnOps(ops, axis)
  for (let step = relevant.length - 1; step >= 0; step -= 1) {
    const op = relevant[step]
    if (!op) continue
    if ('before' in op) {
      position = swapMap(position, op, false)
    } else if (isInsert(op)) {
      if (position >= op.index && position < op.index + op.count) return null
      if (position >= op.index + op.count) position -= op.count
    } else if (position >= op.index) {
      position += op.count
    }
  }
  return position
}

/// Net size change of an axis: screen extent = file extent + delta.
export function netAxisDelta(ops: readonly StructuralOp[], axis: Axis): number {
  let delta = 0
  for (const op of rowColumnOps(ops, axis)) {
    if ('before' in op) continue
    delta += isInsert(op) ? op.count : -op.count
  }
  return delta
}

interface Span {
  start: number
  end: number
}

/// Envelope of a span's image through an insertion: a monotone shift, so the
/// span ends map to the envelope ends.
function insertEnvelope(span: Span, index: number, count: number): Span {
  return {
    start: span.start >= index ? span.start + count : span.start,
    end: span.end >= index ? span.end + count : span.end,
  }
}

/// Envelope of a span's image through a removal; null when the whole span
/// falls inside the removed block.
function removeEnvelope(span: Span, index: number, count: number): Span | null {
  const left = span.start < index ? { start: span.start, end: Math.min(span.end, index - 1) } : null
  const right =
    span.end >= index + count
      ? { start: Math.max(span.start, index + count) - count, end: span.end - count }
      : null
  if (left && right) return { start: left.start, end: right.end }
  return left ?? right
}

/// Envelope of a span's image through one move. The map is a translation on
/// each swap block and the identity elsewhere, so extremes over the span sit
/// at the span ends or at block boundaries inside the span.
function moveEnvelope(
  span: Span,
  op: Extract<RowColumnOp, { before: number }>,
  forward: boolean,
): Span {
  const [a, b] = swapImageBlocks(op, forward)
  let start = Infinity
  let end = -Infinity
  for (const position of [span.start, span.end, a.start, a.end, b.start, b.end]) {
    if (position < span.start || position > span.end) continue
    const mapped = swapMap(position, op, forward)
    start = Math.min(start, mapped)
    end = Math.max(end, mapped)
  }
  return { start, end }
}

/// Envelope of a span's image through the whole op sequence (`forward` =
/// file → screen). Moves make the composite non-monotonic, so the envelope is
/// tracked op by op: each op's blocks are evaluated in the intermediate
/// coordinate space that op actually ran in. Over-reading is safe, tearing is
/// not. Null when nothing in the span survives the mapping.
function spanEnvelope(
  ops: readonly StructuralOp[],
  axis: Axis,
  span: Span,
  forward: boolean,
): Span | null {
  const relevant = rowColumnOps(ops, axis)
  if (!forward) relevant.reverse()
  let current: Span | null = span
  for (const op of relevant) {
    if (!current) return null
    if ('before' in op) {
      current = moveEnvelope(current, op, forward)
    } else if (isInsert(op) === forward) {
      // Inserts applied forward and removals inverted both shift lines apart.
      current = insertEnvelope(current, op.index, op.count)
    } else {
      current = removeEnvelope(current, op.index, op.count)
    }
  }
  return current
}

/// File-space range backing a screen-space range. The result may span file
/// lines that were deleted (they map back to nothing and are dropped on
/// install). Null when no line in the range has file backing.
export function screenRangeToFileRange(
  ops: readonly StructuralOp[],
  range: CellArea,
): CellArea | null {
  const rows = spanEnvelope(ops, 'row', { start: range.startRow, end: range.endRow }, false)
  const columns = spanEnvelope(
    ops,
    'column',
    { start: range.startColumn, end: range.endColumn },
    false,
  )
  if (!rows || !columns) return null
  return {
    startRow: rows.start,
    endRow: rows.end,
    startColumn: columns.start,
    endColumn: columns.end,
  }
}

/// Screen-space extent of a file-space range; null when every line in the
/// range was deleted.
export function fileRangeToScreenRange(
  ops: readonly StructuralOp[],
  range: CellArea,
): CellArea | null {
  const rows = spanEnvelope(ops, 'row', { start: range.startRow, end: range.endRow }, true)
  const columns = spanEnvelope(
    ops,
    'column',
    { start: range.startColumn, end: range.endColumn },
    true,
  )
  if (!rows || !columns) return null
  return {
    startRow: rows.start,
    endRow: rows.end,
    startColumn: columns.start,
    endColumn: columns.end,
  }
}

/// Screen position of the indexing cutoff: the last screen row whose file
/// row is indexed. Screen rows above it are either indexed or inserted.
export function indexedThroughScreenRow(
  ops: readonly StructuralOp[],
  indexedThroughFileRow: number | null,
): number | null {
  if (indexedThroughFileRow === null) return null
  for (let fileRow = indexedThroughFileRow; fileRow >= 0; fileRow -= 1) {
    const screenRow = fileToScreen(ops, 'row', fileRow)
    if (screenRow !== null) return screenRow
  }
  return -1
}

/// Translates a sidecar range result (file coordinates) into screen
/// coordinates. Cells, row properties, and hyperlinks on deleted lines are
/// dropped; merges with a deleted edge are skipped (display-only loss —
/// the save side reshapes merges through the same operation stream).
export function mapRangeResultToScreen(
  ops: readonly StructuralOp[],
  result: WorkbookRangeResult,
): Pick<WorkbookRangeResult, 'cells' | 'rows' | 'merges' | 'hyperlinks'> {
  const cells: WorkbookRangeResult['cells'] = []
  for (const cell of result.cells) {
    const row = fileToScreen(ops, 'row', cell.row)
    const column = fileToScreen(ops, 'column', cell.column)
    if (row === null || column === null) continue
    cells.push({ ...cell, row, column })
  }
  const rows: WorkbookRangeResult['rows'] = []
  for (const rowProperty of result.rows) {
    const row = fileToScreen(ops, 'row', rowProperty.row)
    if (row === null) continue
    rows.push({ ...rowProperty, row })
  }
  const merges: WorkbookRangeResult['merges'] = []
  for (const merge of result.merges) {
    const startRow = fileToScreen(ops, 'row', merge.startRow)
    const endRow = fileToScreen(ops, 'row', merge.endRow)
    const startColumn = fileToScreen(ops, 'column', merge.startColumn)
    const endColumn = fileToScreen(ops, 'column', merge.endColumn)
    if (startRow === null || endRow === null || startColumn === null || endColumn === null) continue
    merges.push({ startRow, endRow, startColumn, endColumn })
  }
  const hyperlinks: WorkbookRangeResult['hyperlinks'] = []
  for (const hyperlink of result.hyperlinks) {
    const row = fileToScreen(ops, 'row', hyperlink.row)
    const column = fileToScreen(ops, 'column', hyperlink.column)
    if (row === null || column === null) continue
    hyperlinks.push({ ...hyperlink, row, column })
  }
  return { cells, rows, merges, hyperlinks }
}
