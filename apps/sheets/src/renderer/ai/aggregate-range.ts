/**
 * aggregate_range tool implementation: batched statistics that never load the
 * grid or touch the formula engine — the supported path for distinct counts
 * and frequency questions on large streamed workbooks. Extracted from App.tsx.
 */
import {
  columnLabel,
  formatAddress,
  parseAddress,
  rangeCellCount,
  type RangeBounds,
} from '../../domain/cell-address'
import type { CellScalar } from '../../domain/workbook.types'
import { isSheetRemoved, journalEntriesInRange } from '../edit-journal'
import { createRangeAggregator, type RangeAggregate } from './aggregate'
import type { WorkbookReadContext } from './workbook-readers'

interface ConstantFill {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly value: CellScalar
}

interface ClippedFill extends ConstantFill {
  readonly priority: number
}

interface FillSegment {
  readonly startColumn: number
  readonly endColumn: number
  readonly value: CellScalar
  readonly priority: number
}

interface FillBand {
  readonly startRow: number
  endRow: number
  readonly segments: FillSegment[]
}

interface SweepEvents {
  readonly add: number[]
  readonly remove: number[]
}

function sweepEvent(events: Map<number, SweepEvents>, position: number): SweepEvents {
  let event = events.get(position)
  if (!event) {
    event = { add: [], remove: [] }
    events.set(position, event)
  }
  return event
}

function heapPush(heap: number[], value: number, fills: readonly ClippedFill[]): void {
  heap.push(value)
  let index = heap.length - 1
  while (index > 0) {
    const parent = Math.floor((index - 1) / 2)
    if (fills[heap[parent]!]!.priority >= fills[value]!.priority) break
    heap[index] = heap[parent]!
    index = parent
  }
  heap[index] = value
}

function heapPop(heap: number[], fills: readonly ClippedFill[]): void {
  const tail = heap.pop()
  if (tail === undefined || heap.length === 0) return
  let index = 0
  while (true) {
    const left = index * 2 + 1
    if (left >= heap.length) break
    const right = left + 1
    const child =
      right < heap.length && fills[heap[right]!]!.priority > fills[heap[left]!]!.priority
        ? right
        : left
    if (fills[heap[child]!]!.priority <= fills[tail]!.priority) break
    heap[index] = heap[child]!
    index = child
  }
  heap[index] = tail
}

function winningColumnSegments(
  fills: readonly ClippedFill[],
  rowActive: ReadonlySet<number>,
): FillSegment[] {
  const events = new Map<number, SweepEvents>()
  for (const fillIndex of rowActive) {
    const fill = fills[fillIndex]!
    sweepEvent(events, fill.startColumn).add.push(fillIndex)
    sweepEvent(events, fill.endColumn + 1).remove.push(fillIndex)
  }
  const points = [...events.keys()].sort((left, right) => left - right)
  const columnActive = new Set<number>()
  const heap: number[] = []
  const segments: FillSegment[] = []
  for (let at = 0; at + 1 < points.length; at += 1) {
    const point = points[at]!
    const event = events.get(point)!
    for (const fillIndex of event.remove) columnActive.delete(fillIndex)
    for (const fillIndex of event.add) {
      columnActive.add(fillIndex)
      heapPush(heap, fillIndex, fills)
    }
    while (heap[0] !== undefined && !columnActive.has(heap[0])) heapPop(heap, fills)
    const winnerIndex = heap[0]
    const nextPoint = points[at + 1]!
    if (winnerIndex === undefined || nextPoint <= point) continue
    const winner = fills[winnerIndex]!
    const previous = segments[segments.length - 1]
    if (previous?.priority === winner.priority && previous.endColumn + 1 === point) {
      segments[segments.length - 1] = { ...previous, endColumn: nextPoint - 1 }
    } else {
      segments.push({
        startColumn: point,
        endColumn: nextPoint - 1,
        value: winner.value,
        priority: winner.priority,
      })
    }
  }
  return segments
}

function sameFillSegments(left: readonly FillSegment[], right: readonly FillSegment[]): boolean {
  return (
    left.length === right.length &&
    left.every(
      (segment, index) =>
        segment.startColumn === right[index]?.startColumn &&
        segment.endColumn === right[index]?.endColumn &&
        segment.priority === right[index]?.priority,
    )
  )
}

function buildFillBands(fills: readonly ConstantFill[], bounds: RangeBounds): FillBand[] {
  const clipped: ClippedFill[] = []
  fills.forEach((fill, priority) => {
    const startRow = Math.max(fill.startRow, bounds.startRow)
    const endRow = Math.min(fill.endRow, bounds.endRow)
    const startColumn = Math.max(fill.startColumn, bounds.startColumn)
    const endColumn = Math.min(fill.endColumn, bounds.endColumn)
    if (startRow > endRow || startColumn > endColumn) return
    clipped.push({
      startRow,
      endRow,
      startColumn,
      endColumn,
      value: fill.value,
      priority,
    })
  })
  if (clipped.length === 0) return []

  const events = new Map<number, SweepEvents>()
  clipped.forEach((fill, fillIndex) => {
    sweepEvent(events, fill.startRow).add.push(fillIndex)
    sweepEvent(events, fill.endRow + 1).remove.push(fillIndex)
  })
  const points = [...events.keys()].sort((left, right) => left - right)
  const rowActive = new Set<number>()
  const bands: FillBand[] = []
  for (let at = 0; at + 1 < points.length; at += 1) {
    const point = points[at]!
    const event = events.get(point)!
    for (const fillIndex of event.remove) rowActive.delete(fillIndex)
    for (const fillIndex of event.add) rowActive.add(fillIndex)
    if (rowActive.size === 0) continue
    const segments = winningColumnSegments(clipped, rowActive)
    if (segments.length === 0) continue
    const endRow = points[at + 1]! - 1
    const previous = bands[bands.length - 1]
    if (
      previous &&
      previous.endRow + 1 === point &&
      sameFillSegments(previous.segments, segments)
    ) {
      previous.endRow = endRow
    } else {
      bands.push({ startRow: point, endRow, segments })
    }
  }
  return bands
}

function fillSegmentAt(
  bands: readonly FillBand[],
  row: number,
  column: number,
): FillSegment | undefined {
  let low = 0
  let high = bands.length - 1
  let band: FillBand | undefined
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    const candidate = bands[middle]!
    if (row < candidate.startRow) high = middle - 1
    else if (row > candidate.endRow) low = middle + 1
    else {
      band = candidate
      break
    }
  }
  if (!band) return undefined
  low = 0
  high = band.segments.length - 1
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    const segment = band.segments[middle]!
    if (column < segment.startColumn) high = middle - 1
    else if (column > segment.endColumn) low = middle + 1
    else return segment
  }
  return undefined
}

function filledAreaInRange(bands: readonly FillBand[], bounds: RangeBounds): number {
  let area = 0
  for (const band of bands) {
    const startRow = Math.max(bounds.startRow, band.startRow)
    const endRow = Math.min(bounds.endRow, band.endRow)
    if (startRow > endRow) continue
    const rowCount = endRow - startRow + 1
    for (const segment of band.segments) {
      const startColumn = Math.max(bounds.startColumn, segment.startColumn)
      const endColumn = Math.min(bounds.endColumn, segment.endColumn)
      if (startColumn <= endColumn) area += rowCount * (endColumn - startColumn + 1)
    }
  }
  return area
}

export async function aggregateWorkbookRange(
  ctx: WorkbookReadContext,
  sheetIdArg: string | undefined,
  bounds: RangeBounds,
): Promise<{ ok: true; aggregate: RangeAggregate } | { ok: false; error: string }> {
  const aggregator = createRangeAggregator()
  const state = ctx.lazyWorkbookRef.current
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const activeSheetId = workbook?.getActiveSheet()?.getSheetId()
  if (!state) {
    // Demo workbook: the snapshot holds every cell in memory.
    const snapshot = ctx.adapterRef.current.getSnapshot()
    const sheet = sheetIdArg
      ? snapshot.sheets.find((candidate) => candidate.id === sheetIdArg)
      : (snapshot.sheets.find((candidate) => candidate.id === activeSheetId) ?? snapshot.sheets[0])
    if (!sheet) {
      return {
        ok: false,
        error: sheetIdArg
          ? `Unknown sheet: ${sheetIdArg} (use an id from get_workbook_context)`
          : 'No workbook is open.',
      }
    }
    const worksheet = workbook?.getSheetBySheetId(sheet.id)
    let counted = 0
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const position = parseAddress(address)
      if (
        position.row < bounds.startRow ||
        position.row > bounds.endRow ||
        position.column < bounds.startColumn ||
        position.column > bounds.endColumn
      ) {
        continue
      }
      let value = cell.value
      // The in-memory model stores value:null for formula cells; the computed
      // value lives in Univer's formula engine — backfill like find_cells.
      if (cell.formula && value === null && worksheet) {
        try {
          value = (worksheet.getRange(address).getValue() as CellScalar) ?? null
        } catch {
          /* fail-open: count the cell as empty */
        }
      }
      aggregator.add(value)
      counted += 1
    }
    aggregator.addEmpty(rangeCellCount(bounds) - counted)
    return { ok: true, aggregate: aggregator.finish(50) }
  }
  const sheetId = sheetIdArg ?? activeSheetId
  if (!sheetId) return { ok: false, error: 'No workbook is open.' }
  if (isSheetRemoved(state.editJournal, sheetId)) {
    return { ok: false, error: `Unknown sheet: ${sheetId}` }
  }
  const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheetMeta) {
    return {
      ok: false,
      error:
        'aggregate_range only works on sheets that came with the file — read sheets added this session with read_range.',
    }
  }
  if ((state.editJournal.structuralOps.get(sheetId) ?? []).length > 0) {
    return {
      ok: false,
      error:
        'aggregate_range is unavailable after row/column structural changes on this sheet in this session — save the file first, then retry.',
    }
  }
  // Cells and declarative fills written this session can sit below/right of
  // the file extent.
  let extentRows = sheetMeta.rowCount
  let extentColumns = sheetMeta.columnCount
  for (const entry of state.editJournal.cells.get(sheetId)?.values() ?? []) {
    if (!entry.hasValue) continue
    extentRows = Math.max(extentRows, entry.row + 1)
    extentColumns = Math.max(extentColumns, entry.column + 1)
  }
  const bulkFills = state.editJournal.bulkConstantFills.get(sheetId) ?? []
  for (const fill of bulkFills) {
    extentRows = Math.max(extentRows, fill.endRow + 1)
    extentColumns = Math.max(extentColumns, fill.endColumn + 1)
  }
  const clamped = {
    startRow: bounds.startRow,
    endRow: Math.min(bounds.endRow, extentRows - 1),
    startColumn: bounds.startColumn,
    endColumn: Math.min(bounds.endColumn, extentColumns - 1),
  }
  if (clamped.startRow > clamped.endRow || clamped.startColumn > clamped.endColumn) {
    return {
      ok: false,
      error: `The range is outside the sheet data extent A1:${columnLabel(extentColumns - 1)}${extentRows}.`,
    }
  }
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  const fillBands = buildFillBands(bulkFills, clamped)
  // Explicit content entries win over every fill; style-only entries do not.
  const journalValues = new Map<string, { row: number; column: number; value: CellScalar }>()
  for (const entry of journalEntriesInRange(state.editJournal, sheetId, clamped)) {
    if (!entry.hasValue) continue
    let value = entry.value ?? null
    // Journal formula entries store value:null; the computed result lives in
    // Univer (journal edits are always applied there) — backfill like find_cells.
    if (entry.formula && value === null && worksheet) {
      try {
        value =
          (worksheet.getRange(formatAddress(entry.row, entry.column)).getValue() as CellScalar) ??
          null
      } catch {
        /* fail-open: count the cell as empty */
      }
    }
    journalValues.set(`${entry.row}:${entry.column}`, {
      row: entry.row,
      column: entry.column,
      value,
    })
  }

  const fillOverrideCounts = new Map<FillSegment, number>()
  for (const entry of journalValues.values()) {
    const segment = fillSegmentAt(fillBands, entry.row, entry.column)
    if (segment) fillOverrideCounts.set(segment, (fillOverrideCounts.get(segment) ?? 0) + 1)
  }
  let counted = 0
  for (const band of fillBands) {
    const rowCount = band.endRow - band.startRow + 1
    for (const segment of band.segments) {
      const repetitions =
        rowCount * (segment.endColumn - segment.startColumn + 1) -
        (fillOverrideCounts.get(segment) ?? 0)
      aggregator.addRepeated(segment.value, repetitions)
      counted += repetitions
    }
  }

  const fileEndRow = Math.min(clamped.endRow, sheetMeta.rowCount - 1)
  const fileEndColumn = Math.min(clamped.endColumn, sheetMeta.columnCount - 1)
  if (clamped.startRow <= fileEndRow && clamped.startColumn <= fileEndColumn) {
    const width = fileEndColumn - clamped.startColumn + 1
    const batchRows = Math.max(1, Math.floor(18_000 / width))
    for (let startRow = clamped.startRow; startRow <= fileEndRow; startRow += batchRows) {
      const endRow = Math.min(startRow + batchRows - 1, fileEndRow)
      const batchBounds = {
        startRow,
        endRow,
        startColumn: clamped.startColumn,
        endColumn: fileEndColumn,
      }
      // Fully overlaid chunks do not depend on the sidecar or its indexing
      // progress. This is the common whole-column-fill path.
      if (filledAreaInRange(fillBands, batchBounds) === rangeCellCount(batchBounds)) continue
      let result
      try {
        result = await window.desktopApi.readWorkbookRange({
          sessionId: state.file.sessionId,
          sheetId,
          range: batchBounds,
        })
      } catch (error: unknown) {
        return {
          ok: false,
          error: error instanceof Error ? error.message : 'Failed to read the range.',
        }
      }
      if (result.indexedThroughRow === null || result.indexedThroughRow < endRow) {
        return {
          ok: false,
          error: 'The range is still being indexed — retry after workbook indexing completes.',
        }
      }
      for (const cell of result.cells) {
        if (
          cell.row < startRow ||
          cell.row > endRow ||
          cell.column < clamped.startColumn ||
          cell.column > fileEndColumn ||
          journalValues.has(`${cell.row}:${cell.column}`) ||
          fillSegmentAt(fillBands, cell.row, cell.column)
        ) {
          continue
        }
        aggregator.add(cell.value)
        counted += 1
      }
    }
  }
  for (const { value } of journalValues.values()) {
    aggregator.add(value)
    counted += 1
  }
  aggregator.addEmpty(rangeCellCount(clamped) - counted)
  return { ok: true, aggregate: aggregator.finish(50) }
}
