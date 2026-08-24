/**
 * What the composer's scope chip shows, and how a raw grid selection becomes
 * one. The chip describes the request the user is waiting on, not the grid:
 * once a run has frozen its scope, clicking around (or dropping the chip) can
 * no longer retarget that run, so a chip that kept tracking the live selection
 * would name a range the run is not using.
 *
 * The raw selection also needs two corrections before it can be shown or sent.
 * Clicking a column header selects the sheet's entire million-row height, which
 * is unreadable on the chip and a nonsense range in the model's context, so
 * everything is clamped to the data extent first. And a column is a name to the
 * user ("Amount"), not a letter, so a selection covering whole columns is
 * labelled from the header row rather than by coordinates.
 */

import { columnLabel, type RangeBounds } from '../../domain/cell-address'
import type { FrozenSelection } from './tools'

export interface ScopeChip {
  /** A1 notation to label the chip with; null shows no chip at all */
  readonly range: string | null
  /** the scope belongs to a run in flight, so it can no longer be dropped */
  readonly locked: boolean
  /** header names when the scope covers whole columns; they label the chip in
   *  place of the range, because that is what the user picked */
  readonly columns?: readonly string[]
}

/**
 * `runScope` is the scope of the run in flight: undefined when no run owns one,
 * null when the user had dropped it before sending. `liveScope` is the current
 * grid selection and `dismissed` whether the user cleared the chip for the next
 * run.
 */
export function resolveScopeChip(
  runScope: FrozenSelection | null | undefined,
  liveScope: FrozenSelection | null,
  dismissed: boolean,
): ScopeChip {
  const scope = runScope !== undefined ? runScope : dismissed ? null : liveScope
  return {
    range: scope?.a1 ?? null,
    locked: runScope !== undefined,
    ...(scope?.columns ? { columns: scope.columns } : {}),
  }
}

/** Last row/column index holding data; both -1 on a sheet with no data. */
export interface DataExtent {
  readonly lastRow: number
  readonly lastColumn: number
}

/** Longest header text kept before an ellipsis, so one verbose header cannot
 *  push the chip over the composer width. */
const MAX_HEADER_CHARS = 24
/** Past this many columns a list of names stops reading as a label. */
const MAX_NAMED_COLUMNS = 3

/**
 * Cap a selection at the sheet's data extent: a whole-column click arrives as a
 * million rows, and neither the user nor the model should be told that is what
 * was selected.
 *
 * Each axis is capped on its own, and only when the selection starts inside the
 * data on that axis. Capping the axes together would abandon both whenever one
 * inverts — a header click on a column past the data would keep its million
 * rows. Leaving an axis alone when the selection starts past the data keeps a
 * deliberately empty block (a spot the user picked to build something in) at
 * the size they marked out.
 */
export function clampBoundsToExtent(bounds: RangeBounds, extent: DataExtent): RangeBounds {
  const cap = (start: number, end: number, last: number): number =>
    start <= last ? Math.min(end, last) : end
  return {
    startRow: bounds.startRow,
    startColumn: bounds.startColumn,
    endRow: cap(bounds.startRow, bounds.endRow, extent.lastRow),
    endColumn: cap(bounds.startColumn, bounds.endColumn, extent.lastColumn),
  }
}

export function boundsToA1(bounds: RangeBounds): string {
  const start = `${columnLabel(bounds.startColumn)}${bounds.startRow + 1}`
  if (bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn) return start
  return `${start}:${columnLabel(bounds.endColumn)}${bounds.endRow + 1}`
}

/**
 * Header names for a selection that covers whole columns, or null when it is
 * not column-shaped. Column-shaped means it starts at row 1 and reaches the
 * last row holding data — the user clicked column headers (or dragged the full
 * height) rather than marking out a block inside the data.
 *
 * Row 1 is taken as the header row the way an Excel table does; a sheet without
 * headers gets labelled by its first row of values, which is still the most
 * recognizable thing in that column.
 */
export function columnScopeHeaders(
  bounds: RangeBounds,
  extent: DataExtent,
  headerAt: (column: number) => string,
): readonly string[] | null {
  if (extent.lastRow < 0) return null
  if (bounds.startRow !== 0 || bounds.endRow < extent.lastRow) return null
  if (bounds.endColumn - bounds.startColumn + 1 > MAX_NAMED_COLUMNS) return null
  const headers: string[] = []
  for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
    const header = headerAt(column).trim()
    // an unnamed column cannot be described any better than by its letter
    if (!header) return null
    headers.push(
      header.length > MAX_HEADER_CHARS ? `${header.slice(0, MAX_HEADER_CHARS)}…` : header,
    )
  }
  return headers
}
