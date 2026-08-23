/// Allow-edit ranges are kept in current screen coordinates on
/// LazyWorkbookState: mapped from file space once when a sheet finishes
/// indexing, then incrementally through each recorded structural op — so the
/// dialog and the save read them without further translation (the save's
/// protectedRanges rewrite runs after structural replay).

import { columnLabel, parseRange } from '../domain/cell-address'
import type { StructuralJournalOp } from './edit-journal'
import { fileRangeToScreenRange, fileToScreen } from './view-transform'

export interface ProtectedRangeEntry {
  readonly name: string
  readonly sqref: string
  readonly hasPassword: boolean
}

function formatArea(area: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): string {
  const start = `${columnLabel(area.startColumn)}${area.startRow + 1}`
  if (area.startRow === area.endRow && area.startColumn === area.endColumn) return start
  return `${start}:${columnLabel(area.endColumn)}${area.endRow + 1}`
}

interface Area {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/// The range envelope over-reads on a move-rows that partially overlaps the
/// area (its span semantics serve viewport fetches) — for an edit whitelist
/// that would fail open. With a move in play, rows are mapped one by one and
/// the survivors split into contiguous runs instead.
const EXACT_MOVE_ROW_CAP = 50_000

function mapArea(area: Area, ops: readonly StructuralJournalOp[]): Area[] {
  const hasMove = ops.some((op) => op.kind === 'move-rows')
  if (!hasMove || area.endRow - area.startRow > EXACT_MOVE_ROW_CAP) {
    const moved = fileRangeToScreenRange(ops, area)
    return moved === null ? [] : [moved]
  }
  let probeRow: number | null = null
  const rows: number[] = []
  for (let row = area.startRow; row <= area.endRow; row += 1) {
    const screen = fileToScreen(ops, 'row', row)
    if (screen === null) continue
    if (probeRow === null) probeRow = row
    rows.push(screen)
  }
  if (probeRow === null) return []
  // Moves never touch columns; the column span still shifts through any
  // insert/remove-cols in the op list.
  const columns = fileRangeToScreenRange(ops, { ...area, startRow: probeRow, endRow: probeRow })
  if (columns === null) return []
  rows.sort((a, b) => a - b)
  const areas: Area[] = []
  for (const row of rows) {
    const last = areas[areas.length - 1]
    if (last && row === last.endRow + 1) {
      last.endRow = row
    } else {
      areas.push({
        startRow: row,
        endRow: row,
        startColumn: columns.startColumn,
        endColumn: columns.endColumn,
      })
    }
  }
  return areas
}

/// Maps a sqref (one or more space-separated A1 areas) through structural
/// ops; fully deleted areas drop out, an unparseable area stays verbatim.
/// null when nothing survives.
function mapSqref(sqref: string, ops: readonly StructuralJournalOp[]): string | null {
  const parts = sqref
    .split(/\s+/)
    .filter((part) => part !== '')
    .flatMap((part) => {
      let area
      try {
        area = parseRange(part.replaceAll('$', ''))
      } catch {
        return [part]
      }
      return mapArea(area, ops).map(formatArea)
    })
  return parts.length === 0 ? null : parts.join(' ')
}

/// Ranges with every area deleted drop out entirely.
export function mapProtectedRanges(
  ranges: readonly ProtectedRangeEntry[],
  ops: readonly StructuralJournalOp[],
): ProtectedRangeEntry[] {
  if (ops.length === 0) return [...ranges]
  const mapped: ProtectedRangeEntry[] = []
  for (const range of ranges) {
    const sqref = mapSqref(range.sqref, ops)
    if (sqref !== null) mapped.push({ ...range, sqref })
  }
  return mapped
}
