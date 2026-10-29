/// The model-facing view of a workbook: every cell the reader can see, with a
/// withheld cell written out as its `{{label}}` instead of its value.
///
/// The read tools in `workbook-readers.ts` each project one read of the grid;
/// this is the single projection for a caller that wants a sheet as one block of
/// text rather than as a series of tool calls. Both ask the same index
/// (`./redact`), so a cell withheld here is withheld there too.
///
/// ## The index is a parameter, and an unusable one withholds everything
///
/// A caller that failed to pass the index would otherwise hand the model every
/// value the reader hid — silently, and in total. That is the same reason
/// `redact-load.ts` resolves an unreadable redaction part to `WITHHELD_UNKNOWN`
/// rather than to "nothing withheld": not knowing is not the same as knowing
/// there is nothing. So the parameter is not optional, and anything that is not
/// a usable index is treated as `WITHHELD_UNKNOWN` rather than as
/// `NO_REDACTIONS`. A sheet with no id is withheld the same way: its cells
/// cannot be looked up, and "no mark found" would not mean "no mark".
///
/// ## Total, and total in the direction that hides
///
/// Nothing here throws. A shape this function does not recognise still produces
/// text, because a throw is worse than a wrong answer here: the caller's
/// fallback is its own unredacted workbook. A cell whose address cannot be
/// placed, and a cell whose value has no shape we recognise, are both written as
/// `{{private}}` — a visible false refusal, which is the better trade against a
/// silent leak.

import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { placeholderSource, type RedactionIndex } from './redact'
import { WITHHELD_UNKNOWN } from './redact-load'

/** The label used when a cell cannot be resolved against the index. */
export const UNKNOWN_LABEL = 'private'

/** One cell, as the file model already holds it (`CellState` satisfies this). */
export interface ModelViewCell {
  readonly value: CellScalar
}

/**
 * One sheet and its cells, keyed by A1 address — the same shape
 * `WorksheetState.cells` and `readCells` speak, so a caller hands over what it
 * already has rather than translating for this.
 */
export interface ModelViewSheet {
  /** the id the read tools address the sheet by; without it nothing can be resolved */
  readonly id: string
  readonly name: string
  readonly cells: Readonly<Record<string, ModelViewCell>>
}

/** The workbook, or the part of it the caller wants handed to the model. */
export interface ModelViewWorkbook {
  readonly sheets: readonly ModelViewSheet[]
}

/** A cell address that was placed on the grid; `row`/`column` are zero-based. */
interface PlacedCell {
  readonly address: string
  readonly row: number
  readonly column: number
}

/**
 * Where an address that is not one is sorted. A large number rather than
 * `Infinity`, because a comparator returning `NaN` leaves the order up to the
 * engine — an output that varies between runs is untestable.
 */
const UNPLACED = Number.MAX_SAFE_INTEGER

/** A cap on the projection: stop after this many cell lines / characters. */
export interface ModelTextLimit {
  /** Non-empty cell lines across all sheets, first sheets first. */
  maxLines: number
  /** Characters of joined output, roughly. */
  maxChars: number
}

/**
 * The text of a workbook as the model may read it.
 *
 * `limit` is for callers that want a sketch rather than the whole workbook —
 * file naming reads one sheet's worth of cells at most. Without it the walk
 * covers every cell of every sheet, which is what the read tools already do
 * and what a chat turn pays for deliberately.
 */
export function modelTextOf(
  workbook: ModelViewWorkbook,
  redactions: RedactionIndex,
  limit?: ModelTextLimit,
): string {
  const index = usableIndex(redactions) ? redactions : WITHHELD_UNKNOWN
  const sheets = Array.isArray(workbook?.sheets) ? workbook.sheets : []
  const blocks: string[] = []
  const budget = limit ? { lines: limit.maxLines, chars: limit.maxChars } : null
  for (const sheet of sheets) {
    if (!sheet || typeof sheet !== 'object') continue
    const lines = sheetLines(sheet, index, budget)
    // An empty sheet still prints its heading: "this sheet has no cells" is
    // information, and dropping it makes a withheld sheet indistinguishable
    // from one the model was never shown.
    blocks.push(
      lines.length === 0
        ? `## ${sheetHeading(sheet)}`
        : `## ${sheetHeading(sheet)}\n${lines.join('\n')}`,
    )
    // The budget is spent: later sheets would only be walked to be dropped.
    if (budget !== null && (budget.lines <= 0 || budget.chars <= 0)) break
  }
  return blocks.join('\n\n')
}

/** What is left of a {@link ModelTextLimit} while the lines are emitted. */
interface TextBudget {
  lines: number
  chars: number
}

/**
 * `sheetLines` for one sheet: one `A1: value` line per cell worth printing.
 *
 * With a budget the walk stops at the first cell that would overshoot it, so
 * a workbook too large to project costs a bounded walk rather than a
 * hundred-thousand-line string nobody reads past the start of.
 */
function sheetLines(
  sheet: ModelViewSheet,
  index: RedactionIndex,
  budget?: TextBudget | null,
): string[] {
  const cells = sheet.cells
  if (!cells || typeof cells !== 'object') return []
  const sheetId = sheet.id
  const out: string[] = []
  for (const cell of placed(cells)) {
    const label = labelAt(index, sheetId, cell)
    // `displayedValue` only returns null for an empty cell, and the empty case
    // is the only one not worth a line — a withheld cell is always worth one,
    // because the mark is exactly what the model is meant to see.
    const text = label === null ? displayedValue(cells[cell.address]) : placeholderSource(label)
    if (text === null) continue
    if (budget) {
      if (budget.lines <= 0 || budget.chars <= 0) break
      budget.lines -= 1
      budget.chars -= cell.address.length + text.length + 2
    }
    out.push(`${cell.address}: ${text}`)
  }
  return out
}

/**
 * Every cell's address in reading order.
 *
 * An address that is not one cannot be asked about, so it keeps the raw key as
 * its address (keys are addresses by contract — echoing one costs nothing) and
 * sorts last, where it is withheld rather than printed.
 */
function placed(cells: Readonly<Record<string, unknown>>): PlacedCell[] {
  const out: PlacedCell[] = []
  for (const address of Object.keys(cells)) {
    try {
      const at = parseAddress(address)
      out.push({ address, row: at.row, column: at.column })
    } catch {
      out.push({ address, row: UNPLACED, column: UNPLACED })
    }
  }
  return out.sort((left, right) => left.row - right.row || left.column - right.column)
}

/**
 * The label covering a cell, or null when it is not withheld.
 *
 * Every way this can fail resolves to `UNKNOWN_LABEL` rather than to null: a
 * label lookup that throws, or a sheet whose id is missing, means the cell was
 * never resolved — and an unresolved cell must not read as a clear one.
 */
function labelAt(index: RedactionIndex, sheetId: unknown, cell: PlacedCell): string | null {
  if (typeof sheetId !== 'string' || sheetId === '') return UNKNOWN_LABEL
  if (cell.row === UNPLACED) return UNKNOWN_LABEL
  try {
    return index.labelAt(sheetId, cell.row, cell.column)
  } catch {
    return UNKNOWN_LABEL
  }
}

/** A cell's displayed value, or null when there is nothing to show. */
function displayedValue(cell: unknown): string | null {
  if (!cell || typeof cell !== 'object') return null
  const value = (cell as { value?: unknown }).value
  if (value === null || value === undefined) return null
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return oneLine(String(value))
  }
  // Not a value we can read. Printing `String(value)` of a shape we do not
  // recognise is how a secret held in an unexpected field gets serialised, so
  // the cell reads as withheld instead.
  return placeholderSource(UNKNOWN_LABEL)
}

/** A cell is one line here, so a value carrying line breaks is folded. */
function oneLine(value: string): string {
  return value.replace(/[\t\r\n]+/g, ' ')
}

/** The heading a sheet is listed under, in the app's own `name (id=…)` form. */
function sheetHeading(sheet: ModelViewSheet): string {
  const name = typeof sheet.name === 'string' && sheet.name !== '' ? sheet.name : '(unnamed)'
  return typeof sheet.id === 'string' && sheet.id !== '' ? `${name} (id=${sheet.id})` : name
}

/** An index is usable when it can answer about a cell at all. */
function usableIndex(index: RedactionIndex): boolean {
  return Boolean(index) && typeof index.labelAt === 'function'
}
