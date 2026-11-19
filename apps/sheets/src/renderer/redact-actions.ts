/**
 * What the grid's right-click item does with the selected cells.
 *
 * One item does both jobs, as in slides: a fresh selection asks what the cells
 * stand for, and a selection that is already withheld stops withholding it.
 * The alternative — a second item — would mean a second string in a locale set
 * that must not grow, and the two would be indistinguishable on a mixed
 * selection.
 *
 * The mark is a label laid over the reader's own cells. Nothing here writes to
 * a cell: a withheld cell keeps its value, keeps its format, and keeps
 * calculating. What changes is only which view the model gets.
 */
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

import { buildRedactionIndex, type RedactionIndex, type SheetRefLike } from './ai/redact'
import type { SelectionRequest } from './redact-menu'

/** The open question the dialog is asking. */
export interface RedactDialogState {
  readonly sheetId: string
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  /** A1 notation of the selection, so the dialog can name what it covers */
  readonly rangeLabel: string
  /** The cell's own text, offered as a starting label; empty for a range */
  readonly seed: string
}

/** What the right-click should do, before any label exists. */
export type RedactIntent =
  | { readonly kind: 'ask'; readonly dialog: RedactDialogState }
  | { readonly kind: 'clear'; readonly sheetName: string; readonly mark: MarkKey }
  /** The click named a sheet this workbook does not have open. */
  | { readonly kind: 'ignore' }

/** Identifies one mark without its label, so a re-mark can replace it. */
export interface MarkKey {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

const COLUMN_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

/** `0 -> A`, `26 -> AA`. Spreadsheets address columns by label, not index. */
export function columnLabel(index: number): string {
  let n = index
  let out = ''
  for (;;) {
    out = COLUMN_LETTERS[n % 26] + out
    n = Math.floor(n / 26) - 1
    if (n < 0) break
  }
  return out
}

function address(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`
}

export function rangeLabelOf(request: SelectionRequest): string {
  const start = address(request.startRow, request.startColumn)
  if (request.isSingleCell) return start
  return `${start}:${address(request.endRow, request.endColumn)}`
}

function sameArea(a: MarkKey, b: MarkKey): boolean {
  return (
    a.startRow === b.startRow &&
    a.endRow === b.endRow &&
    a.startColumn === b.startColumn &&
    a.endColumn === b.endColumn
  )
}

function covers(area: MarkKey, request: SelectionRequest): boolean {
  return (
    area.startRow <= request.startRow &&
    area.endRow >= request.endRow &&
    area.startColumn <= request.startColumn &&
    area.endColumn >= request.endColumn
  )
}

/**
 * Decide what a right-click means.
 *
 * A selection already inside a mark clears it; anything else asks. The check
 * is containment rather than equality so selecting a superset of a marked cell
 * — clicking the cell and dragging past it, the way a spreadsheet selection
 * usually works — still reaches the mark rather than silently creating a
 * second, overlapping one.
 */
export function redactIntentFor(
  request: SelectionRequest,
  index: RedactionIndex,
  sheetNameOf: (sheetId: string) => string | undefined,
): RedactIntent {
  const sheetName = sheetNameOf(request.sheetId)
  if (sheetName === undefined) return { kind: 'ignore' }
  const mark = index.marksFor(request.sheetId).find((candidate) => covers(candidate, request))
  if (mark) {
    return {
      kind: 'clear',
      sheetName,
      mark: {
        startRow: mark.startRow,
        endRow: mark.endRow,
        startColumn: mark.startColumn,
        endColumn: mark.endColumn,
      },
    }
  }
  return {
    kind: 'ask',
    dialog: {
      sheetId: request.sheetId,
      sheetName,
      startRow: request.startRow,
      endRow: request.endRow,
      startColumn: request.startColumn,
      endColumn: request.endColumn,
      rangeLabel: rangeLabelOf(request),
      seed: '',
    },
  }
}

/** Drop `target` from the marks, keeping every other sheet and mark. */
export function clearMark(
  states: readonly SheetRedactionState[],
  sheetName: string,
  target: MarkKey,
): SheetRedactionState[] {
  return states
    .map((state) =>
      state.sheetName === sheetName
        ? { sheetName, marks: state.marks.filter((mark) => !sameArea(mark, target)) }
        : state,
    )
    .filter((state) => state.marks.length > 0)
}

/** Add `mark` to `sheetName`, replacing any mark covering exactly the same area. */
export function addMark(
  states: readonly SheetRedactionState[],
  sheetName: string,
  mark: MarkKey & { readonly label: string; readonly previousFill?: string | null },
): SheetRedactionState[] {
  const without = clearMark(states, sheetName, mark)
  const existing = without.find((state) => state.sheetName === sheetName)
  const next: SheetRedactionState = {
    sheetName,
    marks: [...(existing?.marks ?? []), mark],
  }
  return existing ? without.map((state) => (state === existing ? next : state)) : [...without, next]
}

/**
 * Rebuild the reader index after a change.
 *
 * Rebuilt from the same state the save uses rather than patched, so there is
 * exactly one source of truth: an index that disagreed with the part would
 * withhold cells the file does not record, and the next save would then write
 * the wrong marks.
 */
export function indexFor(
  states: readonly SheetRedactionState[],
  sheets: readonly SheetRefLike[],
): RedactionIndex {
  return buildRedactionIndex(states, sheets)
}

/**
 * True when the reader's marks differ from what the file was opened with.
 *
 * The save ticks decide whether there is anything to write from the edit
 * journal alone, and a mark is not in the journal — so without this a workbook
 * whose only edit is a withheld cell never autosaves, never gets a crash
 * recovery copy, and loses the mark when the tab closes.
 *
 * Compared as serialized text: the marks are a handful of rectangles, and both
 * sides come from the same writer, so the comparison is exact and a mark that
 * moved by one row counts as a change.
 */
export function hasPendingRedactionChange(
  current: readonly SheetRedactionState[],
  loaded: readonly SheetRedactionState[],
): boolean {
  return JSON.stringify(current) !== JSON.stringify(loaded)
}
