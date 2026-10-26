/// Writes a declarative filter snapshot into worksheet XML: the
/// `<autoFilter>` element (with per-column value / custom criteria) plus row
/// visibility inside the filter's row span. Unsupported criteria fail closed.
/// A filter whose range sits inside an Excel table is routed into the table
/// part instead (Excel stores table filters there, never on the worksheet) —
/// see `applyTableFilterState` and `applyRowVisibilityState`.

export class FilterEditError extends Error {}

interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface FilterColumnState {
  readonly colId: number
  readonly values?: readonly string[] | undefined
  readonly blank?: boolean | undefined
  readonly customs?:
    | {
        readonly and?: boolean | undefined
        readonly filters: readonly {
          readonly val: string | number
          readonly operator?: string | undefined
        }[]
      }
    | undefined
}

export interface SheetFilterState {
  readonly sheetName: string
  readonly filter: {
    readonly range: CellArea
    readonly columns: readonly FilterColumnState[]
  } | null
  readonly hiddenRows: readonly number[]
  readonly visibilityRange: CellArea
}

const CUSTOM_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])

export function applyFilterState(worksheetXml: string, state: SheetFilterState): string {
  const element = state.filter === null ? '' : serializeAutoFilter(state.filter)
  const existing = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.exec(
    worksheetXml,
  )
  let xml = worksheetXml
  if (existing) {
    xml =
      worksheetXml.slice(0, existing.index) +
      element +
      worksheetXml.slice(existing.index + existing[0].length)
  } else if (element !== '') {
    xml = insertAfterSheetData(worksheetXml, element)
  }
  return applyRowVisibilityState(xml, state)
}

/// The worksheet half of a table-routed filter: row visibility only. The
/// `<autoFilter>` element itself lives in the table part, and a worksheet
/// element overlapping a table is invalid OOXML that Excel repairs away.
export function applyRowVisibilityState(worksheetXml: string, state: SheetFilterState): string {
  return applyRowVisibility(worksheetXml, state.visibilityRange, new Set(state.hiddenRows))
}

/// Geometry of the table part a filter routed into, in final (post-shift)
/// coordinates: the full table ref including the totals band.
export interface TableFilterRoute {
  readonly area: CellArea
  /// headerRowCount/@headerRowCount of the table part (Excel default 1).
  readonly headerRowCount: number
  /// totalsRowCount/@totalsRowCount of the table part (Excel default 0).
  readonly totalsRowCount: number
}

/// Rewrites the `<autoFilter>` element inside a table part from the
/// declarative snapshot. Excel keeps the table's filter here, spanning the
/// header row through the last data row (the totals band is excluded), with
/// filterColumn/@colId offsets from the table's left edge. `filter: null`
/// (or a snapshot without criteria) leaves the bare element — Clear Filter
/// semantics; a headerless table cannot carry filter buttons, so its element
/// is dropped like the table-add/table-edit paths do.
export function applyTableFilterState(
  tableXml: string,
  route: TableFilterRoute,
  filter: SheetFilterState['filter'],
): string {
  const filterPattern = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/
  if (route.headerRowCount === 0) {
    return tableXml.replace(filterPattern, '')
  }
  const ref = toRef({ ...route.area, endRow: route.area.endRow - route.totalsRowCount })
  let element = `<autoFilter ref="${ref}"/>`
  if (filter !== null) {
    const offset = filter.range.startColumn - route.area.startColumn
    const width = route.area.endColumn - route.area.startColumn + 1
    for (const column of filter.columns) {
      const colId = column.colId + offset
      if (colId < 0 || colId > width - 1) {
        throw new FilterEditError('A filter column falls outside its table and cannot be saved.')
      }
    }
    const columns = [...filter.columns]
      .sort((left, right) => left.colId - right.colId)
      .map((column) => serializeFilterColumn(column, offset))
      .join('')
    if (columns !== '') element = `<autoFilter ref="${ref}">${columns}</autoFilter>`
  }
  if (filterPattern.test(tableXml)) {
    return tableXml.replace(filterPattern, () => element)
  }
  return tableXml.replace(/<table\b[^>]*>/, (open) => `${open}${element}`)
}

function serializeAutoFilter(filter: NonNullable<SheetFilterState['filter']>): string {
  const ref = toRef(filter.range)
  const columns = [...filter.columns]
    .sort((left, right) => left.colId - right.colId)
    .map((column) => serializeFilterColumn(column, 0))
    .join('')
  return columns === ''
    ? `<autoFilter ref="${ref}"/>`
    : `<autoFilter ref="${ref}">${columns}</autoFilter>`
}

/// `columnOffset` re-bases wire colIds (relative to the snapshot range's
/// first column) onto the element's own left edge; worksheet snapshots pass 0.
function serializeFilterColumn(column: FilterColumnState, columnOffset: number): string {
  const colId = column.colId + columnOffset
  if (colId < 0) {
    throw new FilterEditError('A filter column falls outside its range and cannot be saved.')
  }
  const parts: string[] = []
  if (column.values !== undefined || column.blank) {
    const blank = column.blank ? ' blank="1"' : ''
    const values = (column.values ?? [])
      .map((value) => `<filter val="${escapeXmlAttribute(value)}"/>`)
      .join('')
    parts.push(`<filters${blank}>${values}</filters>`)
  }
  if (column.customs) {
    for (const custom of column.customs.filters) {
      if (custom.operator !== undefined && !CUSTOM_OPERATORS.has(custom.operator)) {
        throw new FilterEditError(
          `Filter condition "${custom.operator}" cannot be saved as XLSX yet.`,
        )
      }
    }
    const and = column.customs.and ? ' and="1"' : ''
    const filters = column.customs.filters
      .map((custom) => {
        const operator =
          custom.operator === undefined || custom.operator === 'equal'
            ? ''
            : ` operator="${custom.operator}"`
        return `<customFilter${operator} val="${escapeXmlAttribute(String(custom.val))}"/>`
      })
      .join('')
    parts.push(`<customFilters${and}>${filters}</customFilters>`)
  }
  if (parts.length === 0) return ''
  return `<filterColumn colId="${colId}">${parts.join('')}</filterColumn>`
}

/// Schema order places autoFilter after sheetData (and after the protection
/// block when present).
function insertAfterSheetData(worksheetXml: string, element: string): string {
  let insertAt = -1
  for (const pattern of [
    /<sheetData\s*\/>|<\/sheetData>/,
    /<sheetProtection\b[^>]*\/>/,
    /<protectedRanges\b[^>]*\/>|<\/protectedRanges>/,
    /<\/scenarios>/,
  ]) {
    const match = pattern.exec(worksheetXml)
    if (match) insertAt = Math.max(insertAt, match.index + match[0].length)
  }
  if (insertAt < 0) throw new FilterEditError('Worksheet has no sheetData element.')
  return worksheetXml.slice(0, insertAt) + element + worksheetXml.slice(insertAt)
}

/// Declarative row visibility inside the filter's data rows: listed rows are
/// hidden, every other row in the span is unhidden (matching how Excel
/// re-evaluates a filter). The header row is never touched.
function applyRowVisibility(
  worksheetXml: string,
  range: CellArea,
  hiddenRows: ReadonlySet<number>,
): string {
  const firstDataRow = range.startRow + 1
  const seen = new Set<number>()
  let xml = worksheetXml.replace(
    /<row\b([^>]*?)(\/>|>)/g,
    (full, attributes: string, close: string) => {
      const rowNumber = /(?:^|\s)r="([0-9]+)"/.exec(attributes)?.[1]
      if (rowNumber === undefined) return full
      const rowIndex = Number(rowNumber) - 1
      if (rowIndex < firstDataRow || rowIndex > range.endRow) return full
      seen.add(rowIndex)
      const withoutHidden = attributes.replace(/\s*hidden="[^"]*"/, '')
      const hidden = hiddenRows.has(rowIndex) ? ' hidden="1"' : ''
      return `<row${withoutHidden}${hidden}${close}`
    },
  )
  const missing = [...hiddenRows]
    .filter(
      (rowIndex) => !seen.has(rowIndex) && rowIndex >= firstDataRow && rowIndex <= range.endRow,
    )
    .sort((left, right) => left - right)
  for (const rowIndex of missing) {
    xml = insertEmptyHiddenRow(xml, rowIndex + 1)
  }
  return xml
}

function insertEmptyHiddenRow(worksheetXml: string, rowNumber: number): string {
  const newRow = `<row r="${rowNumber}" hidden="1"/>`
  const rowStartPattern = /<row\b[^>]*?\br="([1-9][0-9]*)"/g
  let match: RegExpExecArray | null
  while ((match = rowStartPattern.exec(worksheetXml)) !== null) {
    if (Number(match[1]) > rowNumber) {
      return worksheetXml.slice(0, match.index) + newRow + worksheetXml.slice(match.index)
    }
  }
  if (worksheetXml.includes('</sheetData>')) {
    return worksheetXml.replace('</sheetData>', () => `${newRow}</sheetData>`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(worksheetXml)) {
    return worksheetXml.replace(emptySheetData, () => `<sheetData>${newRow}</sheetData>`)
  }
  throw new FilterEditError('Worksheet has no sheetData element.')
}

function toRef(range: CellArea): string {
  return (
    `${columnToLetters(range.startColumn)}${range.startRow + 1}` +
    `:${columnToLetters(range.endColumn)}${range.endRow + 1}`
  )
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}
