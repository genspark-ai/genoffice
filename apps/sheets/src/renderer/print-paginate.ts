/// Excel's page planner: each print area starts a new page; rows and
/// columns never split, so a row (column) that does not fit the remaining
/// space moves to the next page; the heading strip and the repeated title
/// rows/columns take their share of every page; a manual break ends the
/// page before the row/column it names. Pages run down then over (or over
/// then down). Pure — the layout feeds dimensions, the result drives both
/// the printed HTML and the preview's page list.

export interface PrintArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface IndexSpan {
  readonly start: number
  readonly end: number
}

export interface PaginationInput {
  readonly areas: readonly PrintArea[]
  /// Print-space size (pt) of one row / column; 0 for hidden ones.
  readonly rowHeightPt: (row: number) => number
  readonly columnWidthPt: (column: number) => number
  readonly titleRows: IndexSpan | null
  readonly titleColumns: IndexSpan | null
  /// 0-based index of the row / column that begins a new page.
  readonly rowBreaks: readonly number[]
  readonly colBreaks: readonly number[]
  /// Content a page holds in print space (printable paper size / scale).
  readonly pageWidthPt: number
  readonly pageHeightPt: number
  /// Row-number strip width / column-letter strip height (0 when headings are off).
  readonly headingWidthPt: number
  readonly headingHeightPt: number
  readonly order?: 'downThenOver' | 'overThenDown'
}

/// One printed page: the body rows/columns it carries (title rows/columns
/// repeat on top of them and are never part of the body).
export interface PrintPage {
  readonly area: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

/// Splits `indices` into bands that fit `capacity`, starting a new band at
/// every manual break. A single index taller than the capacity still gets a
/// band of its own.
export function bandIndices(
  indices: readonly number[],
  sizeOf: (index: number) => number,
  capacity: number,
  breaks: ReadonlySet<number>,
): IndexSpan[] {
  const bands: IndexSpan[] = []
  let start: number | null = null
  let used = 0
  let last = -1
  for (const index of indices) {
    const size = sizeOf(index)
    const open = start !== null
    if (open && (breaks.has(index) || used + size > capacity)) {
      bands.push({ start: start as number, end: last })
      start = null
      used = 0
    }
    if (start === null) start = index
    used += size
    last = index
  }
  if (start !== null) bands.push({ start, end: last })
  return bands
}

function inSpan(span: IndexSpan | null, index: number): boolean {
  return span !== null && index >= span.start && index <= span.end
}

function spanSize(span: IndexSpan | null, sizeOf: (index: number) => number): number {
  if (span === null) return 0
  let total = 0
  for (let index = span.start; index <= span.end; index += 1) total += sizeOf(index)
  return total
}

/// Body rows of an area: hidden rows and the repeated title rows drop out.
export function bodyRows(
  area: PrintArea,
  input: Pick<PaginationInput, 'titleRows' | 'rowHeightPt'>,
): number[] {
  const rows: number[] = []
  for (let row = area.startRow; row <= area.endRow; row += 1) {
    if (inSpan(input.titleRows, row) || input.rowHeightPt(row) <= 0) continue
    rows.push(row)
  }
  return rows
}

export function bodyColumns(
  area: PrintArea,
  input: Pick<PaginationInput, 'titleColumns' | 'columnWidthPt'>,
): number[] {
  const columns: number[] = []
  for (let column = area.startColumn; column <= area.endColumn; column += 1) {
    if (inSpan(input.titleColumns, column) || input.columnWidthPt(column) <= 0) continue
    columns.push(column)
  }
  return columns
}

export function planPrintPages(input: PaginationInput): PrintPage[] {
  const rowCapacity =
    input.pageHeightPt - input.headingHeightPt - spanSize(input.titleRows, input.rowHeightPt)
  const columnCapacity =
    input.pageWidthPt - input.headingWidthPt - spanSize(input.titleColumns, input.columnWidthPt)
  const rowBreaks = new Set(input.rowBreaks)
  const colBreaks = new Set(input.colBreaks)
  const pages: PrintPage[] = []
  input.areas.forEach((area, areaIndex) => {
    const rows = bodyRows(area, input)
    const columns = bodyColumns(area, input)
    if (rows.length === 0 || columns.length === 0) return
    const rowBands = bandIndices(rows, input.rowHeightPt, rowCapacity, rowBreaks)
    const columnBands = bandIndices(columns, input.columnWidthPt, columnCapacity, colBreaks)
    const page = (rowBand: IndexSpan, columnBand: IndexSpan): PrintPage => ({
      area: areaIndex,
      rowStart: rowBand.start,
      rowEnd: rowBand.end,
      colStart: columnBand.start,
      colEnd: columnBand.end,
    })
    if (input.order === 'overThenDown') {
      for (const rowBand of rowBands) {
        for (const columnBand of columnBands) pages.push(page(rowBand, columnBand))
      }
    } else {
      for (const columnBand of columnBands) {
        for (const rowBand of rowBands) pages.push(page(rowBand, columnBand))
      }
    }
  })
  return pages
}
