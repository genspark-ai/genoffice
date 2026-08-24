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

/// Image of a segment through an insertion: lines at or past the index shift
/// apart; a straddling segment splits around the inserted block.
function insertSegments(segment: Span, index: number, count: number): Span[] {
  if (segment.end < index) return [segment]
  if (segment.start >= index) return [{ start: segment.start + count, end: segment.end + count }]
  return [
    { start: segment.start, end: index - 1 },
    { start: index + count, end: segment.end + count },
  ]
}

/// Image of a segment through a removal: the part inside the removed block is
/// gone, survivors past it shift down.
function removeSegments(segment: Span, index: number, count: number): Span[] {
  const out: Span[] = []
  if (segment.start < index) {
    out.push({ start: segment.start, end: Math.min(segment.end, index - 1) })
  }
  if (segment.end >= index + count) {
    out.push({ start: Math.max(segment.start, index + count) - count, end: segment.end - count })
  }
  return out
}

/// Image of a segment through one move: a translation on each swap block and
/// the identity elsewhere, so the segment splits into at most four pieces.
function moveSegments(
  segment: Span,
  op: Extract<RowColumnOp, { before: number }>,
  forward: boolean,
): Span[] {
  const [a, b] = swapImageBlocks(op, forward)
  const aLength = a.end - a.start + 1
  const bLength = b.end - b.start + 1
  const out: Span[] = []
  if (segment.start < a.start) {
    out.push({ start: segment.start, end: Math.min(segment.end, a.start - 1) })
  }
  const inFirst = {
    start: Math.max(segment.start, a.start) + bLength,
    end: Math.min(segment.end, a.end) + bLength,
  }
  if (inFirst.start <= inFirst.end) out.push(inFirst)
  const inSecond = {
    start: Math.max(segment.start, b.start) - aLength,
    end: Math.min(segment.end, b.end) - aLength,
  }
  if (inSecond.start <= inSecond.end) out.push(inSecond)
  if (segment.end > b.end) {
    out.push({ start: Math.max(segment.start, b.end + 1), end: segment.end })
  }
  return out
}

/// Disjoint images of a span through the whole op sequence (`forward` =
/// file → screen). Moves make the composite non-monotonic, so segments are
/// tracked op by op, each op evaluated in the intermediate coordinate space
/// it actually ran in. Tracking exact occupancy (not a bounding box) keeps
/// removals honest: a move that pulls unrelated lines through the tracked
/// area never leaves phantom coverage behind. Empty when nothing in the span
/// survives the mapping.
function spanSegments(
  ops: readonly StructuralOp[],
  axis: Axis,
  span: Span,
  forward: boolean,
): Span[] {
  const relevant = rowColumnOps(ops, axis)
  if (!forward) relevant.reverse()
  let segments: Span[] = [span]
  for (const op of relevant) {
    if (segments.length === 0) break
    if ('before' in op) {
      segments = segments.flatMap((segment) => moveSegments(segment, op, forward))
    } else if (isInsert(op) === forward) {
      // Inserts applied forward and removals inverted both shift lines apart.
      segments = segments.flatMap((segment) => insertSegments(segment, op.index, op.count))
    } else {
      segments = segments.flatMap((segment) => removeSegments(segment, op.index, op.count))
    }
  }
  return segments
}

/// Bounding envelope of the surviving segments; null when none survive. May
/// still span gaps between segments, but both ends sit on real survivors.
function spanEnvelope(
  ops: readonly StructuralOp[],
  axis: Axis,
  span: Span,
  forward: boolean,
): Span | null {
  const segments = spanSegments(ops, axis, span, forward)
  if (segments.length === 0) return null
  let start = Infinity
  let end = -Infinity
  for (const segment of segments) {
    start = Math.min(start, segment.start)
    end = Math.max(end, segment.end)
  }
  return { start, end }
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

/// Exact screen-space images of a file-space range, as disjoint rectangles
/// (surviving row segments × surviving column segments). Unlike the envelope
/// above, the rectangles never cover lines that were deleted or unrelated
/// lines a move shuffled between the survivors. Empty when nothing survives.
export function fileRangeToScreenRanges(ops: readonly StructuralOp[], range: CellArea): CellArea[] {
  const rows = spanSegments(ops, 'row', { start: range.startRow, end: range.endRow }, true)
  const columns = spanSegments(
    ops,
    'column',
    { start: range.startColumn, end: range.endColumn },
    true,
  )
  return rows.flatMap((row) =>
    columns.map((column) => ({
      startRow: row.start,
      endRow: row.end,
      startColumn: column.start,
      endColumn: column.end,
    })),
  )
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
