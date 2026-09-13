/// PivotChart data assembly: turns a pivot definition + the output area's current
/// values into the categories/series needed for chart creation. series/cat point
/// at output-area cells (absolute A1 references), so after a pivot refresh writes
/// the output area back, the chart in Excel follows the references.
///
/// Trade-offs aligned with Excel's PivotChart semantics:
///  - Only data rows/columns (t="data") are used; subtotals and grand totals do
///    not enter the chart (Excel doesn't draw them either);
///  - Cell references are only written when the data rows are contiguous in the
///    output area, otherwise we fall back to literal data (multi-level layouts
///    with interleaved subtotal rows cannot be expressed as one range reference);
///  - TODO(OOXML): wiring chartSpace <c:pivotSource> + pivotFmts (so Excel
///    recognizes the chart as a PivotChart and shows field buttons) is heavy;
///    for now we ship a regular chart + output-area references, which still
///    follows refreshes when opened in Excel.

import { columnLabel } from './cell-address'
import type { PivotDefinition, PivotLayoutLine } from '../gateway/xlsx-pivot'

/// Chart point cap: oversized pivots are truncated to the first N data rows/cols
/// (neither the literal cache nor rendering can handle tens of thousands of
/// points; references only cover the included rows, staying consistent after
/// truncation).
const MAX_PIVOT_CHART_ROWS = 500
const MAX_PIVOT_CHART_SERIES = 20

export interface PivotChartBounds {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

export interface PivotChartSeries {
  readonly name: string
  readonly categories: readonly string[]
  readonly values: readonly number[]
  readonly valuesRef?: string | undefined
  readonly categoriesRef?: string | undefined
}

export interface PivotChartData {
  readonly title: string
  readonly series: readonly PivotChartSeries[]
  /// True when truncated (for user hints); references/data remain consistent.
  readonly truncated: boolean
}

type OutputCell = string | number | boolean | null | undefined

/// Display caption of a member on a layout row/column (sharedItems value; empty
/// members display as an empty string).
function memberCaption(definition: PivotDefinition, field: number, member: number): string {
  const item = definition.fieldItems[field]?.[member]
  if (!item || item.x === null) return ''
  return String(definition.fields[field]?.sharedItems[item.x] ?? '')
}

/// Caption of one data layout line: fixed member captions per level joined with
/// " - "; the data-field level (-2) writes the data-field name (distinguishing
/// series in multi-value pivots).
function dataLineLabel(
  definition: PivotDefinition,
  axisFields: readonly number[],
  line: PivotLayoutLine,
): string {
  const parts: string[] = []
  for (let level = 0; level < axisFields.length; level += 1) {
    const field = axisFields[level]
    const member = line.members[level]
    if (field === undefined || member === null || member === undefined) continue
    parts.push(
      field === -2
        ? (definition.dataFields[member]?.name ?? '')
        : memberCaption(definition, field, member),
    )
  }
  return parts.filter((part) => part !== '').join(' - ')
}

function toNumber(value: OutputCell): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const numeric = Number(String(value ?? '').trim())
  return String(value ?? '').trim() !== '' && Number.isFinite(numeric) ? numeric : 0
}

/// Absolute A1 reference for one column (0-based row numbers); the sheet name is
/// always quoted (matching the existing reference style of chart creation; Excel
/// tolerates redundant quotes).
function columnRef(sheetName: string, column: number, fromRow: number, toRow: number): string {
  const col = columnLabel(column)
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${col}$${fromRow + 1}:$${col}$${toRow + 1}`
}

/// Assembles chart data from the pivot definition + current output-area values.
/// bounds is the 0-based area of outputRef; outputValues are the area's complete
/// values (header row / label columns included).
export function buildPivotChartData(
  definition: PivotDefinition,
  sheetName: string,
  bounds: PivotChartBounds,
  outputValues: readonly (readonly OutputCell[])[],
): PivotChartData {
  // Layout indices of data rows/columns (subtotal/grand-total/blank lines are
  // all excluded).
  const allDataRows = definition.rowLines.flatMap((line, index) =>
    line.t === 'data' ? [index] : [],
  )
  const allDataCols = definition.colLines.flatMap((line, index) =>
    line.t === 'data' ? [index] : [],
  )
  const truncated =
    allDataRows.length > MAX_PIVOT_CHART_ROWS || allDataCols.length > MAX_PIVOT_CHART_SERIES
  const dataRows = allDataRows.slice(0, MAX_PIVOT_CHART_ROWS)
  const dataCols = allDataCols.slice(0, MAX_PIVOT_CHART_SERIES)

  if (dataRows.length === 0 || dataCols.length === 0) {
    return { title: definition.dataFields[0]?.name ?? 'PivotChart', series: [], truncated }
  }

  // A single range reference is only possible when the data rows form one
  // contiguous run in the layout (a trailing grand total doesn't matter);
  // multi-level layouts with interleaved subtotal rows fall back to literal data.
  const rowsContiguous = dataRows.every((rowIndex, i) => rowIndex === dataRows[0]! + i)
  const firstOutRow = bounds.startRow + definition.firstDataRow + (dataRows[0] ?? 0)
  const lastOutRow =
    bounds.startRow + definition.firstDataRow + (dataRows[dataRows.length - 1] ?? 0)

  const categories = dataRows.map((rowIndex) => {
    const label = dataLineLabel(definition, definition.rowFields, definition.rowLines[rowIndex]!)
    return label === '' ? 'Total' : label
  })
  // Categories references are only written with a single label column:
  // multi-level row labels spread across columns and cannot be expressed as one.
  const categoriesRef =
    rowsContiguous && definition.firstDataCol === 1
      ? columnRef(sheetName, bounds.startColumn, firstOutRow, lastOutRow)
      : undefined

  const series: PivotChartSeries[] = dataCols.map((colIndex, seriesIndex) => {
    const line = definition.colLines[colIndex]!
    const label = dataLineLabel(definition, definition.colFields, line)
    const name =
      label !== ''
        ? label
        : (definition.dataFields[line.dataField]?.name ?? `Series ${seriesIndex + 1}`)
    const outColumn = bounds.startColumn + definition.firstDataCol + colIndex
    const values = dataRows.map((rowIndex) =>
      toNumber(
        outputValues[definition.firstDataRow + rowIndex]?.[definition.firstDataCol + colIndex],
      ),
    )
    return {
      name,
      categories,
      values,
      ...(rowsContiguous
        ? { valuesRef: columnRef(sheetName, outColumn, firstOutRow, lastOutRow) }
        : {}),
      ...(categoriesRef !== undefined ? { categoriesRef } : {}),
    }
  })

  const title = definition.dataFields.length === 1 ? definition.dataFields[0]!.name : 'PivotChart'
  return { title, series, truncated }
}
