/// Data → Consolidate: aggregates several source areas into a target range,
/// either by position or by left-column labels. Output cells are live
/// formulas referencing the sources, so they journal and save like any edit.

import { columnLabel, parseRange } from '../domain/cell-address'
import { t } from './i18n/locale'

export type ConsolidateFn = 'sum' | 'count' | 'average' | 'max' | 'min'

export interface ConsolidateConfig {
  readonly fn: ConsolidateFn
  readonly references: readonly string[]
  readonly leftLabels: boolean
}

export interface ConsolidateArea {
  readonly sheetName: string
  readonly startRow: number
  readonly startColumn: number
  readonly rows: number
  readonly columns: number
}

const REFERENCE_PATTERN =
  /^(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)$/

/// "B2:D9", "Sheet2!A1:C4", or "'My Sheet'!A1" — qualifier optional.
export function parseConsolidateReference(
  text: string,
): { sheetName: string | null; range: ReturnType<typeof parseRange> } | null {
  const match = REFERENCE_PATTERN.exec(text.trim())
  if (!match) return null
  const qualifier = match[1]
  const sheetName =
    qualifier === undefined
      ? null
      : qualifier.startsWith("'")
        ? qualifier.slice(1, -1).replaceAll("''", "'")
        : qualifier
  try {
    return { sheetName, range: parseRange(match[2]?.toUpperCase().replaceAll('$', '') ?? '') }
  } catch {
    return null
  }
}

export function isValidConsolidateReference(text: string): boolean {
  return parseConsolidateReference(text) !== null
}

function sheetPrefix(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? `${name}!` : `'${name.replaceAll("'", "''")}'!`
}

function cellRef(area: ConsolidateArea, row: number, column: number): string {
  return `${sheetPrefix(area.sheetName)}${columnLabel(area.startColumn + column)}${area.startRow + row + 1}`
}

function columnRangeRef(area: ConsolidateArea, column: number, absolute: boolean): string {
  const letter = (absolute ? '$' : '') + columnLabel(area.startColumn + column)
  const first = `${letter}${absolute ? '$' : ''}${area.startRow + 1}`
  const last = `${letter}${absolute ? '$' : ''}${area.startRow + area.rows}`
  return `${sheetPrefix(area.sheetName)}${first}:${last}`
}

const POSITION_FN: Record<ConsolidateFn, string> = {
  sum: 'SUM',
  count: 'COUNTA',
  average: 'AVERAGE',
  max: 'MAX',
  min: 'MIN',
}

export interface OutputCell {
  readonly v?: string | null
  readonly f?: string
}

/// By-position: output extent is the max extent across areas; each output
/// cell aggregates the same offset from every area that covers it.
export function buildPositionMatrix(
  fn: ConsolidateFn,
  areas: readonly ConsolidateArea[],
): OutputCell[][] {
  const rows = Math.max(...areas.map((area) => area.rows))
  const columns = Math.max(...areas.map((area) => area.columns))
  const matrix: OutputCell[][] = []
  for (let row = 0; row < rows; row += 1) {
    const line: OutputCell[] = []
    for (let column = 0; column < columns; column += 1) {
      const refs = areas
        .filter((area) => row < area.rows && column < area.columns)
        .map((area) => cellRef(area, row, column))
      line.push(refs.length === 0 ? { v: null } : { f: `=${POSITION_FN[fn]}(${refs.join(',')})` })
    }
    matrix.push(line)
  }
  return matrix
}

/// By left-column labels: rows are matched by their first-column value, in
/// order of first appearance across areas. `areaLabels` holds each area's
/// label column values, aligned with `areas`. Value columns keep their
/// by-position offsets (column 1..n of each area).
export function buildLabelMatrix(
  fn: ConsolidateFn,
  areas: readonly ConsolidateArea[],
  areaLabels: readonly (readonly string[])[],
  target: { row: number; column: number },
): OutputCell[][] | string {
  if (fn === 'max' || fn === 'min') {
    return t('appConsolidateMaxMinLabels')
  }
  if (areas.some((area) => area.columns < 2)) {
    return t('appConsolidateLabelNeedsCols')
  }
  const labels: string[] = []
  const seen = new Set<string>()
  for (const areaLabelList of areaLabels) {
    for (const label of areaLabelList) {
      if (label === '' || seen.has(label)) continue
      seen.add(label)
      labels.push(label)
    }
  }
  if (labels.length === 0) return t('appConsolidateNoLabels')
  const valueColumns = Math.max(...areas.map((area) => area.columns)) - 1
  const matrix: OutputCell[][] = []
  labels.forEach((label, index) => {
    const criteria = `$${columnLabel(target.column)}${target.row + index + 1}`
    const line: OutputCell[] = [{ v: label }]
    for (let column = 0; column < valueColumns; column += 1) {
      const parts = areas
        .map((area, areaIndex) => ({
          area,
          hasLabel: areaLabels[areaIndex]?.includes(label) ?? false,
        }))
        .filter(({ area, hasLabel }) => hasLabel && column + 1 < area.columns)
        .map(({ area }) => ({
          labelRange: columnRangeRef(area, 0, true),
          valueRange: columnRangeRef(area, column + 1, false),
        }))
      if (parts.length === 0) {
        line.push({ v: null })
        continue
      }
      const sumIfs = parts.map((part) => `SUMIF(${part.labelRange},${criteria},${part.valueRange})`)
      const countIfs = parts.map((part) => `COUNTIF(${part.labelRange},${criteria})`)
      const formula =
        fn === 'sum'
          ? sumIfs.join('+')
          : fn === 'count'
            ? countIfs.join('+')
            : `(${sumIfs.join('+')})/(${countIfs.join('+')})`
      line.push({ f: `=${formula}` })
    }
    matrix.push(line)
  })
  return matrix
}

/// Excel refuses to consolidate into a range that overlaps a source area.
export function targetOverlapsSource(
  areas: readonly ConsolidateArea[],
  activeSheetName: string,
  target: { row: number; column: number; rows: number; columns: number },
): boolean {
  return areas.some(
    (area) =>
      area.sheetName === activeSheetName &&
      target.row < area.startRow + area.rows &&
      area.startRow < target.row + target.rows &&
      target.column < area.startColumn + area.columns &&
      area.startColumn < target.column + target.columns,
  )
}
