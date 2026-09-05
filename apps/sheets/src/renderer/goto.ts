/// Reference resolution for the Name Box and the Go To dialog, kept pure so
/// it unit-tests without Univer. Turns user input (an A1 address like `C5`,
/// a range like `B2:D5`, or a defined name like `SalesData`) into an
/// A1-notation string that FWorksheet.getRange() accepts — Univer's parser
/// tolerates `$` markers and `Sheet1!` prefixes but silently yields NaN rows
/// for garbage, so validity is decided here, not by try/catch alone.

import { columnIndex } from '../domain/cell-address'

export interface GoToNameEntry {
  readonly name: string
  readonly ref: string
}

const CELL_OR_RANGE = /^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/
const COLUMN_SPAN = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/
const ROW_SPAN = /^\$?\d+:\$?\d+$/

/// Splits an optional `Sheet1!` / `'My Sheet'!` prefix from the reference
/// body. Input without `!` is all body.
function splitSheetPrefix(ref: string): { readonly sheet: string; readonly body: string } {
  const bang = ref.lastIndexOf('!')
  if (bang === -1) return { sheet: '', body: ref }
  return { sheet: ref.slice(0, bang), body: ref.slice(bang + 1) }
}

/// Excel grid limits: rows 1..1048576, columns A..XFD.
const MAX_ROW_NUMBER = 1048576
const MAX_COLUMN_INDEX = 16383 // XFD

const CELL_PART = /^\$?([A-Za-z]{1,3})\$?(\d+)$/

function columnInBounds(letters: string): boolean {
  try {
    return columnIndex(letters.toUpperCase()) <= MAX_COLUMN_INDEX
  } catch {
    return false
  }
}

function rowInBounds(digits: string): boolean {
  const row = Number(digits)
  return Number.isInteger(row) && row >= 1 && row <= MAX_ROW_NUMBER
}

/// True when every cell endpoint in the shape-validated body also sits
/// inside the grid. Without this, `A0` / `ZZZ1` / `A1048577` / `0:5`
/// pass the shape regexes and leak into getRange as NaN/out-of-bounds,
/// surfacing as a generic failure instead of an unresolvable reference.
function endpointsInBounds(body: string): boolean {
  return body.split(':').every((part) => {
    const cell = CELL_PART.exec(part)
    if (cell?.[1] !== undefined && cell[2] !== undefined) {
      return columnInBounds(cell[1]) && rowInBounds(cell[2])
    }
    const lettersOnly = /^\$?([A-Za-z]{1,3})$/.exec(part)
    if (lettersOnly?.[1] !== undefined) return columnInBounds(lettersOnly[1])
    const digitsOnly = /^\$?(\d+)$/.exec(part)
    if (digitsOnly?.[1] !== undefined) return rowInBounds(digitsOnly[1])
    return false
  })
}

/// True when the reference (after any sheet prefix) is a cell, a range, a
/// whole-column span, or a whole-row span in A1 notation.
export function isA1Reference(ref: string): boolean {
  const { sheet, body } = splitSheetPrefix(ref)
  if (ref.includes('!') && sheet === '') return false
  if (!CELL_OR_RANGE.test(body) && !COLUMN_SPAN.test(body) && !ROW_SPAN.test(body)) {
    return false
  }
  return endpointsInBounds(body)
}

/// Resolves Name Box / Go To input to the A1 reference to jump to, or null
/// when it is neither a valid address nor a resolvable defined name. Defined
/// names win over addresses (names that look like cell refs are invalid
/// anyway) and match case-insensitively. A name whose ref is a
/// formula (e.g. OFFSET) is not a jump target and resolves to null.
export function resolveGoToRef(
  input: string,
  definedNames: readonly GoToNameEntry[],
): string | null {
  const trimmed = input.trim()
  if (trimmed === '') return null
  const lower = trimmed.toLowerCase()
  const named = definedNames.find((entry) => entry.name.toLowerCase() === lower)
  if (named) {
    const ref = named.ref.replace(/^=/, '').trim()
    return isA1Reference(ref) ? ref : null
  }
  return isA1Reference(trimmed) ? trimmed : null
}
