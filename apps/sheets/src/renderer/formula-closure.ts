/// Closure-mode analysis for streamed workbooks: given every formula cell,
/// compute the set of cells the Univer engine needs (formulas plus their
/// referenced precedents) so a large, formula-light workbook can recalculate
/// live without a full preload.

import { FORMULA_REFERENCE_PATTERN, qualifierMatches } from '../gateway/xlsx-structure'
import { swapPosition, toSwapSpans } from './edit-journal'

export interface ClosureSheetInput {
  readonly id: string
  readonly name: string
  readonly rowCount: number
  readonly columnCount: number
  readonly formulas: readonly { row: number; column: number; formula: string }[]
}

export type ClosureResult =
  | {
      readonly ok: true
      /// Per sheet id: every closure cell (formulas and precedents).
      readonly cellsBySheet: Map<string, Set<number>>
      readonly formulaCount: number
      readonly totalCells: number
    }
  | { readonly ok: false; readonly reason: string }

const MAX_COLUMNS = 16_384

export function cellKey(row: number, column: number): number {
  return row * MAX_COLUMNS + column
}

export function keyRow(key: number): number {
  return Math.floor(key / MAX_COLUMNS)
}

export function keyColumn(key: number): number {
  return key % MAX_COLUMNS
}

interface ParsedToken {
  readonly startRow: number | null
  readonly endRow: number | null
  readonly startColumn: number | null
  readonly endColumn: number | null
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

/// Null bounds mean "whole axis" (A:A or 1:5 references), clamped to the
/// target sheet's used range by the caller.
function parseReferenceToken(token: string): ParsedToken | null {
  const wholeColumn = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/.exec(token)
  if (wholeColumn) {
    const start = lettersToColumn(wholeColumn[1] ?? 'A')
    const end = lettersToColumn(wholeColumn[2] ?? 'A')
    return {
      startRow: null,
      endRow: null,
      startColumn: Math.min(start, end),
      endColumn: Math.max(start, end),
    }
  }
  const wholeRow = /^\$?([0-9]+):\$?([0-9]+)$/.exec(token)
  if (wholeRow) {
    const start = Number(wholeRow[1]) - 1
    const end = Number(wholeRow[2]) - 1
    return {
      startRow: Math.min(start, end),
      endRow: Math.max(start, end),
      startColumn: null,
      endColumn: null,
    }
  }
  const cells = token.split(':').map((part) => /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(part))
  if (cells.some((cell) => cell === null)) return null
  const rows = cells.map((cell) => Number(cell?.[2]) - 1)
  const columns = cells.map((cell) => lettersToColumn(cell?.[1] ?? 'A'))
  return {
    startRow: Math.min(...rows),
    endRow: Math.max(...rows),
    startColumn: Math.min(...columns),
    endColumn: Math.max(...columns),
  }
}

interface FormulaReference {
  readonly qualifier: string | undefined
  readonly token: ParsedToken
}

export function parseFormulaReferences(formula: string): FormulaReference[] {
  const references: FormulaReference[] = []
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    for (const match of (segments[index] ?? '').matchAll(FORMULA_REFERENCE_PATTERN)) {
      const token = parseReferenceToken(match[3] ?? '')
      if (token) references.push({ qualifier: match[2], token })
    }
  }
  return references
}

/// True when the formula uses an identifier that is not a function call —
/// a defined name or an external reference the closure cannot resolve.
export function containsUnresolvedNames(formula: string): boolean {
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    // Strip references (and their qualifiers) so ref letters don't register.
    const stripped = (segments[index] ?? '').replace(
      FORMULA_REFERENCE_PATTERN,
      (_full, lead: string) => `${lead} `,
    )
    for (const match of stripped.matchAll(/[A-Za-z_][A-Za-z0-9_.]*/g)) {
      const before = match.index === 0 ? '' : (stripped[match.index - 1] ?? '')
      if (/[A-Za-z0-9_.$'!]/.test(before)) continue
      const name = match[0]
      if (name === 'TRUE' || name === 'FALSE') continue
      const rest = stripped.slice(match.index + name.length)
      if (!/^\s*\(/.test(rest)) return true
    }
  }
  return false
}

export function computeFormulaClosure(
  sheets: readonly ClosureSheetInput[],
  maxCells: number,
): ClosureResult {
  const byName = new Map(sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]))
  const cellsBySheet = new Map<string, Set<number>>()
  let totalCells = 0
  let formulaCount = 0
  const add = (sheetId: string, key: number): boolean => {
    let cells = cellsBySheet.get(sheetId)
    if (!cells) {
      cells = new Set()
      cellsBySheet.set(sheetId, cells)
    }
    if (!cells.has(key)) {
      cells.add(key)
      totalCells += 1
    }
    return totalCells <= maxCells
  }

  for (const sheet of sheets) {
    for (const cell of sheet.formulas) {
      formulaCount += 1
      if (!add(sheet.id, cellKey(cell.row, cell.column))) {
        return { ok: false, reason: `closure exceeds ${maxCells} cells` }
      }
    }
  }
  for (const sheet of sheets) {
    for (const cell of sheet.formulas) {
      if (containsUnresolvedNames(cell.formula)) {
        return { ok: false, reason: 'formula uses a defined name or external reference' }
      }
      for (const reference of parseFormulaReferences(cell.formula)) {
        const target =
          reference.qualifier === undefined
            ? sheet
            : (sheets.find((candidate) =>
                qualifierMatches(reference.qualifier ?? '', candidate.name),
              ) ?? byName.get(unquote(reference.qualifier).toLowerCase()))
        if (!target) {
          return {
            ok: false,
            reason: `formula references an unknown sheet (${reference.qualifier ?? ''})`,
          }
        }
        const startRow = Math.max(reference.token.startRow ?? 0, 0)
        const endRow = Math.min(reference.token.endRow ?? target.rowCount - 1, target.rowCount - 1)
        const startColumn = Math.max(reference.token.startColumn ?? 0, 0)
        const endColumn = Math.min(
          reference.token.endColumn ?? target.columnCount - 1,
          target.columnCount - 1,
        )
        if ((endRow - startRow + 1) * (endColumn - startColumn + 1) > maxCells) {
          return { ok: false, reason: `closure exceeds ${maxCells} cells` }
        }
        for (let row = startRow; row <= endRow; row += 1) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            if (!add(target.id, cellKey(row, column))) {
              return { ok: false, reason: `closure exceeds ${maxCells} cells` }
            }
          }
        }
      }
    }
  }
  return { ok: true, cellsBySheet, formulaCount, totalCells }
}

function unquote(qualifier: string): string {
  return qualifier.startsWith("'") ? qualifier.slice(1, -1).replaceAll("''", "'") : qualifier
}

export interface ClosureRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// Groups closure cells into fetchable row bands, each within the per-read
/// cell budget. Bands merge nearby rows (gap ≤ 32) sharing column bounds.
export function closureFetchRanges(
  cells: ReadonlySet<number>,
  maxCellsPerRange = 18_000,
): ClosureRange[] {
  const columnsByRow = new Map<number, { min: number; max: number }>()
  for (const key of cells) {
    const row = keyRow(key)
    const column = keyColumn(key)
    const bounds = columnsByRow.get(row)
    if (!bounds) columnsByRow.set(row, { min: column, max: column })
    else {
      bounds.min = Math.min(bounds.min, column)
      bounds.max = Math.max(bounds.max, column)
    }
  }
  const rows = [...columnsByRow.keys()].sort((left, right) => left - right)
  const ranges: ClosureRange[] = []
  let band: { startRow: number; endRow: number; minCol: number; maxCol: number } | null = null
  const flush = (): void => {
    if (!band) return
    ranges.push({
      startRow: band.startRow,
      endRow: band.endRow,
      startColumn: band.minCol,
      endColumn: band.maxCol,
    })
    band = null
  }
  for (const row of rows) {
    const bounds = columnsByRow.get(row)
    if (!bounds) continue
    if (band) {
      const minCol = Math.min(band.minCol, bounds.min)
      const maxCol = Math.max(band.maxCol, bounds.max)
      const area = (row - band.startRow + 1) * (maxCol - minCol + 1)
      if (row - band.endRow <= 32 && area <= maxCellsPerRange) {
        band.endRow = row
        band.minCol = minCol
        band.maxCol = maxCol
        continue
      }
      flush()
    }
    band = { startRow: row, endRow: row, minCol: bounds.min, maxCol: bounds.max }
  }
  flush()
  return ranges
}

/// Read bands over formula cells for the sidecar recalc fallback: nearest
/// the viewport first, within the per-request cell budget. Bands past the
/// budget wait for a later run.
export function recalcReadRanges(
  keys: ReadonlySet<number>,
  viewportStartRow: number,
  budget: number,
): ClosureRange[] {
  const bands = closureFetchRanges(keys).sort(
    (left, right) =>
      Math.abs(left.startRow - viewportStartRow) - Math.abs(right.startRow - viewportStartRow),
  )
  const ranges: ClosureRange[] = []
  let used = 0
  for (const band of bands) {
    const area = (band.endRow - band.startRow + 1) * (band.endColumn - band.startColumn + 1)
    if (used + area > budget) continue
    used += area
    ranges.push(band)
    if (ranges.length >= 200) break
  }
  return ranges
}

/// Shifts a screen-space pinned-cell map through one journaled row/column
/// operation (mirrors the edit journal's own entry shifting).
export function shiftPinnedCells<T>(
  entries: ReadonlyMap<string, T>,
  op: { kind: string; index: number; count: number; before?: number },
): Map<string, T> {
  const axis = op.kind === 'insert-cols' || op.kind === 'remove-cols' ? 'column' : 'row'
  const removing = op.kind === 'remove-rows' || op.kind === 'remove-cols'
  const swap =
    op.kind === 'move-rows' && op.before !== undefined
      ? toSwapSpans({ index: op.index, count: op.count, before: op.before })
      : null
  const shifted = new Map<string, T>()
  for (const [key, value] of entries) {
    const [rowText, columnText] = key.split(':')
    const row = Number(rowText)
    const column = Number(columnText)
    const position = axis === 'row' ? row : column
    let moved: number | null
    if (swap) {
      moved = swapPosition(position, swap)
    } else if (removing) {
      if (position >= op.index && position < op.index + op.count) moved = null
      else moved = position >= op.index + op.count ? position - op.count : position
    } else {
      moved = position >= op.index ? position + op.count : position
    }
    if (moved === null) continue
    const next = axis === 'row' ? `${moved}:${column}` : `${row}:${moved}`
    shifted.set(next, value)
  }
  return shifted
}
