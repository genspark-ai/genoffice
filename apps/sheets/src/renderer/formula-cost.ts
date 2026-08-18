/// Static cost analysis for user/AI formulas before they reach Univer's
/// formula engine, which runs on the renderer main thread. Criteria-family
/// functions (COUNTIF/SUMIF/…) and lookups evaluate once per element when
/// their criteria/lookup-value argument is itself a range — the classic
/// distinct-count idiom SUMPRODUCT(1/COUNTIF(D2:D88588,D2:D88588)) does
/// |range|×|criteria| ≈ 7.8e9 comparisons and freezes the app for minutes.
/// Such formulas are rejected at propose/edit time instead.

import { offsetFormulaRefs } from '../domain/formula-shift'
import { qualifierMatches } from '../gateway/xlsx-structure'

/// Like the gateway's FORMULA_REFERENCE_PATTERN, but the unquoted sheet
/// qualifier accepts any Unicode letter — AI/user formulas reference CJK
/// sheet names unquoted (a CJK sheet name before !D2:D88588), exactly the incident shape.
const REFERENCE_PATTERN = new RegExp(
  "(^|[^\\p{L}\\p{N}_.$'!:])" +
    "(?:('(?:[^']|'')+'|[\\p{L}_][\\p{L}\\p{N}_.]*)!)?" +
    '(\\$?[A-Z]{1,3}\\$?[0-9]+(?::\\$?[A-Z]{1,3}\\$?[0-9]+)?' +
    '|\\$?[A-Z]{1,3}:\\$?[A-Z]{1,3}' +
    '|\\$?[0-9]+:\\$?[0-9]+)' +
    '(?![0-9\\p{L}_(])',
  'gu',
)

export interface FormulaCostSheet {
  readonly name: string
  /** data extent; whole-axis references clamp to these */
  readonly rows: number
  readonly columns: number
}

/// Excel grid maximums — the clamp when a whole-axis reference targets a
/// sheet we know nothing about (conservative: unknown big beats unknown ok).
const MAX_SHEET_ROWS = 1_048_576
const MAX_SHEET_COLUMNS = 16_384

/// Above this many element operations for one formula, evaluation visibly
/// blocks the UI thread (measured: 88587² comparisons froze the renderer for
/// minutes; ~4e6 stays well under a second).
export const MAX_QUADRATIC_COST = 4_000_000

/// Per function: [scanArg, perElementArg] index pairs (0-based). The cost of
/// a pair is |scanArg| × |perElementArg| when the per-element argument is
/// itself a multi-cell range (array context: one full scan per element).
/// Repeated pairs (COUNTIFS/SUMIFS families) are expanded dynamically.
const CRITERIA_PAIRS: Readonly<Record<string, 'if' | 'ifs-lead' | 'lookup'>> = {
  COUNTIF: 'if',
  SUMIF: 'if',
  AVERAGEIF: 'if',
  COUNTIFS: 'ifs-lead',
  SUMIFS: 'ifs-lead',
  AVERAGEIFS: 'ifs-lead',
  MINIFS: 'ifs-lead',
  MAXIFS: 'ifs-lead',
  MATCH: 'lookup',
  XMATCH: 'lookup',
  LOOKUP: 'lookup',
  VLOOKUP: 'lookup',
  HLOOKUP: 'lookup',
  XLOOKUP: 'lookup',
}

function argPairs(kind: 'if' | 'ifs-lead' | 'lookup', argCount: number): [number, number][] {
  if (kind === 'if') return [[0, 1]]
  if (kind === 'lookup') return [[1, 0]]
  // COUNTIFS(cr1,c1,cr2,c2,…) — pairs from 0; SUMIFS/… lead with the
  // sum/min/max range, pairs from 1. COUNTIFS is the only 'ifs' without a
  // lead range, and its even/odd layout matches pairs-from-0.
  const pairs: [number, number][] = []
  const start = argCount % 2 === 0 ? 0 : 1
  for (let index = start; index + 1 < argCount; index += 2) {
    pairs.push([index, index + 1])
  }
  return pairs
}

/// Largest referenced-cell count inside one expression fragment; 0 when the
/// fragment contains no reference (scalar criteria — no array context).
function largestReferenceCells(
  fragment: string,
  hostSheet: string,
  sheets: readonly FormulaCostSheet[],
): number {
  let largest = 0
  const segments = fragment.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    for (const match of (segments[index] ?? '').matchAll(REFERENCE_PATTERN)) {
      const qualifier = match[2]
      const token = match[3] ?? ''
      const targetName =
        qualifier === undefined
          ? hostSheet
          : (sheets.find((sheet) => qualifierMatches(qualifier, sheet.name))?.name ?? '')
      const target = sheets.find((sheet) => sheet.name === targetName)
      const rows = target?.rows ?? MAX_SHEET_ROWS
      const columns = target?.columns ?? MAX_SHEET_COLUMNS
      largest = Math.max(largest, referenceCellCount(token, rows, columns))
    }
  }
  return largest
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

function referenceCellCount(token: string, sheetRows: number, sheetColumns: number): number {
  const wholeColumn = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/.exec(token)
  if (wholeColumn) {
    const start = lettersToColumn(wholeColumn[1] ?? 'A')
    const end = lettersToColumn(wholeColumn[2] ?? 'A')
    return (Math.abs(end - start) + 1) * sheetRows
  }
  const wholeRow = /^\$?([0-9]+):\$?([0-9]+)$/.exec(token)
  if (wholeRow) {
    const start = Number(wholeRow[1])
    const end = Number(wholeRow[2])
    return (Math.abs(end - start) + 1) * sheetColumns
  }
  const cells = token.split(':').map((part) => /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(part))
  if (cells.some((cell) => cell === null)) return 0
  const rows = cells.map((cell) => Number(cell?.[2]))
  const columns = cells.map((cell) => lettersToColumn(cell?.[1] ?? 'A'))
  const rowCount = Math.abs((rows[rows.length - 1] ?? 0) - (rows[0] ?? 0)) + 1
  const columnCount = Math.abs((columns[columns.length - 1] ?? 0) - (columns[0] ?? 0)) + 1
  return rowCount * columnCount
}

/// Splits the argument list of the function call whose opening paren sits at
/// `openParen`, honouring nested parens, double-quoted strings, quoted sheet
/// names, and array literals. Returns null on unbalanced input.
function splitArguments(formula: string, openParen: number): string[] | null {
  const args: string[] = []
  let depth = 1
  let braceDepth = 0
  let argStart = openParen + 1
  let index = openParen + 1
  while (index < formula.length) {
    const character = formula[index]
    if (character === '"' || character === "'") {
      const closer = formula.indexOf(character, index + 1)
      if (closer === -1) return null
      index = closer + 1
      continue
    }
    if (character === '{') braceDepth += 1
    else if (character === '}') braceDepth = Math.max(0, braceDepth - 1)
    else if (character === '(') depth += 1
    else if (character === ')') {
      depth -= 1
      if (depth === 0) {
        args.push(formula.slice(argStart, index))
        return args
      }
    } else if (character === ',' && depth === 1 && braceDepth === 0) {
      args.push(formula.slice(argStart, index))
      argStart = index + 1
    }
    index += 1
  }
  return null
}

/// Estimated element operations from array-context criteria/lookup calls in
/// the formula. 0 means "no quadratic pattern detected".
export function estimateQuadraticCost(
  formula: string,
  hostSheet: string,
  sheets: readonly FormulaCostSheet[],
): number {
  let cost = 0
  const pattern = /[A-Za-z_][A-Za-z0-9_.]*(?=\()/g
  for (const match of formula.matchAll(pattern)) {
    // Skip names inside string literals (odd number of quotes before them).
    const before = formula.slice(0, match.index)
    if ((before.match(/"/g)?.length ?? 0) % 2 === 1) continue
    const preceding = match.index === 0 ? '' : (formula[match.index - 1] ?? '')
    if (/[A-Za-z0-9_.$!]/.test(preceding)) continue
    const kind = CRITERIA_PAIRS[match[0].toUpperCase()]
    if (!kind) continue
    const args = splitArguments(formula, match.index + match[0].length)
    if (!args) continue
    for (const [scanIndex, elementIndex] of argPairs(kind, args.length)) {
      const scanFragment = args[scanIndex]
      const elementFragment = args[elementIndex]
      if (scanFragment === undefined || elementFragment === undefined) continue
      const elements = largestReferenceCells(elementFragment, hostSheet, sheets)
      if (elements <= 1) continue
      const scan = Math.max(1, largestReferenceCells(scanFragment, hostSheet, sheets))
      cost += scan * elements
    }
  }
  return cost
}

/// Every formula text inside a set-range-values command payload (`value` /
/// `cellValue` in any of Univer's accepted shapes: single ICellData, 2D
/// array, or {row:{col:…}} matrix — all walked generically).
export function collectCellFormulaTexts(payload: unknown, out: string[] = []): string[] {
  if (!payload || typeof payload !== 'object') return out
  const record = payload as Record<string, unknown>
  if (typeof record.f === 'string' && record.f.startsWith('=')) out.push(record.f)
  for (const child of Object.values(record)) {
    if (child && typeof child === 'object') collectCellFormulaTexts(child, out)
  }
  return out
}

/// File-open degradation: cells whose formula is over the quadratic budget
/// keep only their cached value (like cache-only defined-name cells) so the
/// engine never evaluates them. Returns the input array when nothing
/// degrades. Display-side only — the unedited file formula survives saves.
export function degradeQuadraticFormulaCells<
  T extends { readonly formula?: string | undefined; readonly arrayRef?: string | undefined },
>(cells: T[], hostSheet: string, sheets: readonly FormulaCostSheet[]): T[] {
  let degraded: T[] | null = null
  for (let index = 0; index < cells.length; index += 1) {
    const cell = cells[index]
    if (!cell?.formula) continue
    if (estimateQuadraticCost(cell.formula, hostSheet, sheets) <= MAX_QUADRATIC_COST) continue
    degraded ??= [...cells]
    const { formula: _formula, arrayRef: _arrayRef, ...kept } = cell
    degraded[index] = kept as unknown as T
  }
  return degraded ?? cells
}

/// Model/user-facing guard: null when the formula is fine, otherwise an
/// explanation of why it is rejected.
export function quadraticFormulaError(
  formula: string,
  hostSheet: string,
  sheets: readonly FormulaCostSheet[],
): string | null {
  const cost = estimateQuadraticCost(formula, hostSheet, sheets)
  if (cost <= MAX_QUADRATIC_COST) return null
  return (
    `This formula would perform about ${cost.toLocaleString('en-US')} element comparisons ` +
    '(a criteria/lookup function receives a large range as its per-element argument) and would freeze the app. ' +
    'Do not write distinct-count/array-criteria formulas over large ranges. ' +
    'For statistics questions (distinct counts, frequencies, sums), use the aggregate_range tool and answer in text instead.'
  )
}

/// Fill guard: a filled formula evaluates once per copy, so its
/// per-evaluation cost (largest referenced range plus any quadratic criteria
/// cost) multiplies by the copy count. Relative references that only touch a
/// cell's own row (=B2+1) cost ~1 per copy and pass; absolute ranges over
/// the whole column (=SUM(B$2:B$88588)) or lookups against large tables
/// re-scan everything per row and are rejected. Anchored-start/relative-end
/// ranges (running totals like =SUM(B$2:B2)) expand as they fill, so the
/// copy at the fill's far corner is costed too — the source copy alone
/// would look like one cell per copy and slip through.
export function fillFormulaCostError(
  formula: string,
  copies: number,
  rowDelta: number,
  columnDelta: number,
  hostSheet: string,
  sheets: readonly FormulaCostSheet[],
): string | null {
  const lastCopy = offsetFormulaRefs(formula, rowDelta, columnDelta)
  const candidates = lastCopy === formula ? [formula] : [formula, lastCopy]
  const perEvaluation = Math.max(
    ...candidates.map(
      (candidate) =>
        Math.max(1, largestReferenceCells(candidate, hostSheet, sheets)) +
        estimateQuadraticCost(candidate, hostSheet, sheets),
    ),
  )
  const total = perEvaluation * copies
  if (total <= MAX_QUADRATIC_COST) return null
  return (
    `Filling this formula across ${copies.toLocaleString('en-US')} cells would perform about ` +
    `${total.toLocaleString('en-US')} element operations and freeze the app ` +
    '(each copy re-scans the large range it references). ' +
    'Rewrite it with relative references so each row only touches its own cells, ' +
    'or narrow the referenced range before filling.'
  )
}
