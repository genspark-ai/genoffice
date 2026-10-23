/// Withholding cell values from the model, on the reading side.
///
/// The marks themselves live in `xl/gxRedactions.json` (see
/// `xlsx-gateway/gateway/xlsx-redaction`). This module turns them into the only
/// thing the read tools need: a question they can ask about a cell — "is this
/// withheld, and what is it called?" — answered without ever touching the value.
///
/// Two shapes of leak matter here, and only the first exists in the other apps:
///
/// 1. **Direct.** `read_range` / `read_cells` hand the value over. Fixed by
///    substituting the placeholder.
/// 2. **By arithmetic.** `aggregate_range` computes an exact `sum` over a
///    column. One withheld cell inside it is solvable: sum − everything else.
///    No amount of masking the cell fixes that, because the cell is not the
///    only source of its value. So a withheld cell must not *participate* in a
///    statistic at all — see `addWithheld` in ./aggregate.
///
/// The part is keyed by sheet NAME (that is what the save pipeline addresses),
/// while the read tools address sheets by id. `buildRedactionIndex` is where
/// that translation happens, so no call site has to remember it.

import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

import { parseAddressParts } from '../formula-values'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import {
  placeholderInstruction as buildInstruction,
  sanitizeLabel as coreSanitizeLabel,
  type PlaceholderSpec,
} from '@genoffice/agent-core/redact-core'

/**
 * Clean a label for storage and for the model prompt.
 *
 * The character set is the stricter of the two carriers used across the apps
 * (html's, which also drops `*` and `/` because a label there opens a script
 * comment). A label carries no such risk in a JSON part, but a reader who has
 * learned the rule in one app should not find it relaxed in another.
 */
/**
 * A cell's label, cleaned for a comment carrier.
 *
 * Two characters more than the shared version strips, and deliberately bound
 * here rather than re-exported: a sheet's label is written into an XML
 * comment, where `*` and `/` would close it early and take the rest of the
 * marker with it. A document's label is not in a comment, and stripping these
 * there would rename a span the reader can see.
 */
/** The form the model sees; braces make a marker recognisable in a reply. */
const OPEN = '{{'
const CLOSE = '}}'

export const sanitizeLabel = (raw: string): string => coreSanitizeLabel(raw, { forComment: true })

/**
 * The literal text the model is shown in place of a withheld value.
 *
 * Not the shared one: that builds from the default sanitiser, and a sheet's
 * label goes through the stricter one. Re-exporting it here would put a `*` or
 * a `/` back into a comment carrier — silently, and only for this editor.
 */
export function placeholderSource(label: string): string {
  return `${OPEN}${sanitizeLabel(label) || 'private'}${CLOSE}`
}

/** One withheld rectangle, in zero-based screen coordinates. */
export interface CellMark {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly label: string
}

/** What the read tools ask about a cell. */
export interface RedactionIndex {
  /** the label covering this cell, or null when the cell is not withheld */
  labelAt(sheetId: string, row: number, column: number): string | null
  /** true when this workbook withholds nothing — the common case */
  readonly isEmpty: boolean
  /** the labels in play on one sheet, for the prompt that tells the model */
  labelsFor(sheetId: string): readonly string[]
  /** the raw rectangles, for callers that count an area rather than ask about a cell */
  marksFor(sheetId: string): readonly CellMark[]
  /** every sheet id carrying marks, so callers can decide what to say about them */
  readonly sheetIds: readonly string[]
}

/** A sheet as the app already lists it: the pair the index is built from. */
export interface SheetRefLike {
  readonly id: string
  readonly name: string
}

const EMPTY_INDEX: RedactionIndex = {
  labelAt: () => null,
  isEmpty: true,
  labelsFor: () => [],
  marksFor: () => [],
  sheetIds: [],
}

/**
 * A cell's value as the model may read it.
 *
 * The index is the filter, and this is where a reader goes through it. A cell
 * the reader withheld keeps its real value on the sheet — the point of the
 * feature is that their data survives — so any reader that formats a value for
 * the model has to ask, and forgetting to ask is indistinguishable from never
 * having withheld anything.
 */
export function modelCellValue(
  index: RedactionIndex,
  sheetId: string,
  row: number,
  column: number,
  value: CellScalar,
): CellScalar {
  const label = index.labelAt(sheetId, row, column)
  return label === null ? value : placeholderSource(label)
}

/**
 * The same, from an A1 address, with the sheet the reader is on.
 *
 * Most readers only ever hold an address string, and re-deriving the row and
 * column at each call site is where a second parser — and a second set of
 * edges — comes from.
 */
export function modelCellValueAt(
  index: RedactionIndex,
  sheetId: string | undefined,
  address: string,
  value: CellScalar,
): CellScalar {
  if (!sheetId || index.isEmpty) return value
  const at = parseAddressParts(address)
  if (!at) return value
  return modelCellValue(index, sheetId, at.row, at.column, value)
}

/**
 * The "nothing is withheld" index, shared.
 *
 * Most workbooks never hide a cell, and the readers ask about every cell they
 * touch. Handing them one immutable empty index keeps the per-cell cost at a
 * single property read instead of a null check on every call.
 */
export const NO_REDACTIONS: RedactionIndex = EMPTY_INDEX

function covers(mark: CellMark, row: number, column: number): boolean {
  return (
    row >= mark.startRow &&
    row <= mark.endRow &&
    column >= mark.startColumn &&
    column <= mark.endColumn
  )
}

/**
 * Build the lookup from the marks on disk.
 *
 * A mark naming a sheet this workbook does not have is dropped rather than
 * guessed at: the part is keyed by name, and applying "Sheet1"'s marks to
 * whatever sheet happens to sit at that index would withhold the wrong cells —
 * a silent, wrong-direction error. The reader's own marks are still saved (the
 * save path keys off the same state), so a renamed sheet recovers on the save
 * that renames it.
 */
export function buildRedactionIndex(
  states: readonly SheetRedactionState[],
  sheets: readonly SheetRefLike[],
): RedactionIndex {
  const idByName = new Map(sheets.map((sheet) => [sheet.name, sheet.id]))
  const bySheet = new Map<string, CellMark[]>()
  for (const state of states) {
    const sheetId = idByName.get(state.sheetName)
    if (sheetId === undefined) continue
    const marks = bySheet.get(sheetId) ?? []
    for (const mark of state.marks) {
      marks.push({
        startRow: mark.startRow,
        endRow: mark.endRow,
        startColumn: mark.startColumn,
        endColumn: mark.endColumn,
        label: sanitizeLabel(mark.label),
      })
    }
    bySheet.set(sheetId, marks)
  }
  if (bySheet.size === 0) return EMPTY_INDEX

  return {
    isEmpty: false,
    sheetIds: [...bySheet.keys()],
    labelAt(sheetId, row, column) {
      for (const mark of bySheet.get(sheetId) ?? []) {
        if (covers(mark, row, column)) return mark.label
      }
      return null
    },
    labelsFor(sheetId) {
      const labels = (bySheet.get(sheetId) ?? []).map((mark) => mark.label)
      return [...new Set(labels)]
    },
    marksFor(sheetId) {
      return bySheet.get(sheetId) ?? []
    },
  }
}

/**
 * The number of withheld cells inside a rectangle — the union of the marks,
 * not their sum.
 *
 * Marks routinely overlap: a reader who hides a column and then hides one more
 * cell inside it has two marks covering one cell. Summing the intersections
 * would count that cell twice, and the caller subtracts this from a fill band's
 * repetition count — an over-count there drops real cells out of the range.
 * A scanline over the covered rows, merging the column intervals on each, is
 * exact and stays cheap because the loop is bounded by the marks' own extent.
 */
export function countWithheldIn(
  marks: readonly CellMark[],
  bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): number {
  let firstRow = Number.POSITIVE_INFINITY
  let lastRow = Number.NEGATIVE_INFINITY
  for (const mark of marks) {
    if (mark.endRow < bounds.startRow || mark.startRow > bounds.endRow) continue
    if (mark.endColumn < bounds.startColumn || mark.startColumn > bounds.endColumn) continue
    firstRow = Math.min(firstRow, Math.max(mark.startRow, bounds.startRow))
    lastRow = Math.max(lastRow, Math.min(mark.endRow, bounds.endRow))
  }
  if (firstRow > lastRow) return 0

  const intervals: [number, number][] = []
  let total = 0
  for (let row = firstRow; row <= lastRow; row += 1) {
    intervals.length = 0
    for (const mark of marks) {
      if (row < mark.startRow || row > mark.endRow) continue
      const start = Math.max(mark.startColumn, bounds.startColumn)
      const end = Math.min(mark.endColumn, bounds.endColumn)
      if (end >= start) intervals.push([start, end])
    }
    if (intervals.length === 0) continue
    intervals.sort((left, right) => left[0] - right[0])
    const [firstStart, firstEnd] = intervals[0]!
    let start = firstStart
    let end = firstEnd
    for (let i = 1; i < intervals.length; i += 1) {
      const [nextStart, nextEnd] = intervals[i]!
      if (nextStart <= end + 1) {
        end = Math.max(end, nextEnd)
      } else {
        total += end - start + 1
        start = nextStart
        end = nextEnd
      }
    }
    total += end - start + 1
  }
  return total
}

/**
 * The prompt section that tells the model what the placeholders are.
 *
 * Two things here are specific to a spreadsheet and would be wrong to copy
 * from the other apps:
 *
 * 1. A withheld cell is **excluded from the arithmetic**, not merely masked in
 *    the output. Without saying so, `sum` over a column will not match the
 *    visible values and the model will "correct" it — by guessing the hidden
 *    number, which is the exact thing the reader withheld.
 * 2. A withheld cell does not appear in `find_cells` results at all, so the
 *    model must not conclude the value is absent from the workbook.
 */
/** What a workbook's markers need that a document's or a deck's do not. */
const SHEET_SPEC: PlaceholderSpec = {
  subject: 'workbook',
  opening:
    'Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
  // a sheet's marker cannot be split across cells, not across lines
  boundary: 'a cell boundary',
  middle: [
    'Write text around them as if each stood for the value it replaces, so a row reading "call {{client phone}}" still means what it says.',
    'If a request needs what a placeholder hides, work around it rather than guessing.',
    'Two consequences you must respect, or you will recover what the reader hid:',
    '- Withheld cells are **left out of every statistic**. A sum, average, min, max, distinct count or top-value list over a range that contains one covers only the remaining cells, and `aggregate_range` reports how many were withheld. Never add the withheld values back in, and never treat a total as covering the whole range.',
    '- Withheld cells are **absent from `find_cells` results**. If a search returns nothing, that does not mean the value is not in the workbook — it may simply be withheld. Never conclude a value is missing from the file.',
  ],
}

export function placeholderInstruction(labels: readonly string[]): string {
  return buildInstruction(labels, SHEET_SPEC)
}

export {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  readPlaceholderLabel,
  type PlaceholderIssue,
} from '@genoffice/agent-core/redact-core'
