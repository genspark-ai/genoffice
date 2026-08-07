/// Structural row/column transforms for worksheet XML. Operations replay in
/// order; each rewrites row/cell addresses, same-sheet formula references,
/// merges, and range-scoped features (hyperlinks, autoFilter, data
/// validations, conditional formatting). Anything that cannot be shifted
/// safely throws — the save must fail closed rather than corrupt references.

export type StructuralOp =
  | {
      readonly kind: 'insert-rows' | 'remove-rows' | 'insert-cols' | 'remove-cols'
      readonly index: number
      readonly count: number
    }
  /// Whole-row move: rows [index, index+count) relocate to sit before the
  /// pre-move row `before`. A bijection (nothing is deleted), so row
  /// attributes travel with their rows and references never turn into #REF!.
  | {
      readonly kind: 'move-rows'
      readonly index: number
      readonly count: number
      readonly before: number
    }
  | { readonly kind: 'merge-cells' | 'unmerge-cells'; readonly range: CellArea }
  /// size: points for rows, character width for columns; null = sheet default.
  | {
      readonly kind: 'set-row-size' | 'set-col-size'
      readonly start: number
      readonly end: number
      readonly size: number | null
    }
  | {
      readonly kind: 'set-rows-hidden' | 'set-cols-hidden'
      readonly start: number
      readonly end: number
      readonly hidden: boolean
    }
  /// level: absolute outline level 0-7 (0 removes the attribute); an omitted
  /// collapsed leaves the file's collapsed flag untouched.
  | {
      readonly kind: 'set-rows-outline' | 'set-cols-outline'
      readonly start: number
      readonly end: number
      readonly level: number
      readonly collapsed?: boolean
    }

export type AxisAttributeOp = Extract<StructuralOp, { start: number }>

/// True for operations that shift coordinates or reshape merges — the ones
/// whose sibling parts (drawing anchors, table ranges, calcChain) must move
/// with the sheet. Sizing and visibility changes don't move anything.
export function isShiftingOp(op: StructuralOp): boolean {
  return 'index' in op || 'range' in op
}

interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export type RowColumnOp = Extract<StructuralOp, { index: number }>

export class StructuralShiftError extends Error {}

export function applyStructuralOps(
  worksheetXml: string,
  ops: readonly StructuralOp[],
  sheetName: string,
): string {
  let xml = worksheetXml
  let outlineTouched = false
  for (const op of ops) {
    if ('start' in op) {
      outlineTouched ||= 'level' in op
      xml =
        op.kind === 'set-row-size' ||
        op.kind === 'set-rows-hidden' ||
        op.kind === 'set-rows-outline'
          ? applyRowAttributeOp(xml, op)
          : applyColAttributeOp(xml, op)
      continue
    }
    if (!('index' in op)) {
      xml = applyMergeOp(xml, op)
      continue
    }
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    xml = axis === 'row' ? transformSheetRows(xml, shift) : transformSheetColumns(xml, shift)
    xml = transformFormulas(xml, sheetName, shift, axis)
    xml = transformRangedFeatures(xml, shift, axis)
    if (axis === 'column') xml = transformColDefinitions(xml, shift)
  }
  if (outlineTouched) xml = syncOutlineSummaryLevels(xml)
  return xml
}

/// Excel sizes the outline gutter from sheetFormatPr's outlineLevelRow/Col;
/// recompute both maxima after outline edits so the bars actually show.
function syncOutlineSummaryLevels(xml: string): string {
  let maxRow = 0
  for (const match of xml.matchAll(/<row\b[^>]*?\boutlineLevel="([0-9]+)"/g)) {
    maxRow = Math.max(maxRow, Number(match[1]))
  }
  let maxCol = 0
  for (const match of xml.matchAll(/<col\b[^>]*?\boutlineLevel="([0-9]+)"/g)) {
    maxCol = Math.max(maxCol, Number(match[1]))
  }
  const attributes =
    (maxRow > 0 ? ` outlineLevelRow="${maxRow}"` : '') +
    (maxCol > 0 ? ` outlineLevelCol="${maxCol}"` : '')
  const existing = /<sheetFormatPr\b([^>]*?)(\/?)>/.exec(xml)
  if (existing) {
    const kept = existing[1]!
      .replace(/\s+outlineLevelRow="[^"]*"/g, '')
      .replace(/\s+outlineLevelCol="[^"]*"/g, '')
    return (
      xml.slice(0, existing.index) +
      `<sheetFormatPr${kept}${attributes}${existing[2]}>` +
      xml.slice(existing.index + existing[0].length)
    )
  }
  if (!attributes) return xml
  const anchor = /<cols\b|<sheetData\b/.exec(xml)
  if (!anchor) return xml
  return (
    xml.slice(0, anchor.index) +
    `<sheetFormatPr defaultRowHeight="15"${attributes}/>` +
    xml.slice(anchor.index)
  )
}

function toRef(range: CellArea): string {
  return (
    `${columnToLetters(range.startColumn)}${range.startRow + 1}` +
    `:${columnToLetters(range.endColumn)}${range.endRow + 1}`
  )
}

function applyMergeOp(xml: string, op: Extract<StructuralOp, { range: CellArea }>): string {
  const ref = toRef(op.range)
  if (op.kind === 'unmerge-cells') {
    const without = xml.replace(new RegExp(`<mergeCell\\b[^>]*\\bref="${ref}"[^>]*/>`), '')
    return refreshMergeCount(without)
  }
  const element = `<mergeCell ref="${ref}"/>`
  if (xml.includes('</mergeCells>')) {
    return refreshMergeCount(xml.replace('</mergeCells>', () => `${element}</mergeCells>`))
  }
  // Schema order places mergeCells after autoFilter (when present), which in
  // turn follows sheetData.
  const section = `<mergeCells count="1">${element}</mergeCells>`
  const autoFilter = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.exec(xml)
  if (autoFilter) {
    const insertAt = autoFilter.index + autoFilter[0].length
    return xml.slice(0, insertAt) + section + xml.slice(insertAt)
  }
  if (xml.includes('</sheetData>')) {
    return xml.replace('</sheetData>', () => `</sheetData>${section}`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(xml)) {
    return xml.replace(emptySheetData, (match) => `${match}${section}`)
  }
  throw new StructuralShiftError('Worksheet has no sheetData element.')
}

function formatSize(size: number): string {
  return String(Math.round(size * 100) / 100)
}

/// Sets `ht`/`customHeight` or `hidden` on every row in the span. Rows the
/// file never materialized are created in order — but only when the op sets
/// something (a default-reset on a missing row is already the default).
function applyRowAttributeOp(xml: string, op: AxisAttributeOp): string {
  const seen = new Set<number>()
  let result = xml.replace(/<row\b([^>]*?)(\/>|>)/g, (full, attributes: string, close: string) => {
    const rowNumber = /(?:^|\s)r="([0-9]+)"/.exec(attributes)?.[1]
    if (rowNumber === undefined) return full
    const rowIndex = Number(rowNumber) - 1
    if (rowIndex < op.start || rowIndex > op.end) return full
    seen.add(rowIndex)
    let patched = attributes
    if ('size' in op) {
      patched = patched.replace(/\s*ht="[^"]*"/, '').replace(/\s*customHeight="[^"]*"/, '')
      if (op.size !== null) patched += ` ht="${formatSize(op.size)}" customHeight="1"`
    } else if ('level' in op) {
      patched = patched.replace(/\s*outlineLevel="[^"]*"/, '')
      if (op.level > 0) patched += ` outlineLevel="${op.level}"`
      if (op.collapsed !== undefined) {
        patched = patched.replace(/\s*collapsed="[^"]*"/, '')
        if (op.collapsed) patched += ' collapsed="1"'
      }
    } else {
      patched = patched.replace(/\s*hidden="[^"]*"/, '')
      if (op.hidden) patched += ' hidden="1"'
    }
    return `<row${patched}${close}`
  })
  const setsSomething =
    'size' in op
      ? op.size !== null
      : 'level' in op
        ? op.level > 0 || op.collapsed === true
        : op.hidden
  if (!setsSomething) return result
  const newAttributes =
    'size' in op
      ? ` ht="${formatSize(op.size ?? 0)}" customHeight="1"`
      : 'level' in op
        ? (op.level > 0 ? ` outlineLevel="${op.level}"` : '') +
          (op.collapsed ? ' collapsed="1"' : '')
        : ' hidden="1"'
  for (let rowIndex = op.start; rowIndex <= op.end; rowIndex += 1) {
    if (seen.has(rowIndex)) continue
    result = insertEmptyRow(result, rowIndex + 1, newAttributes)
  }
  return result
}

function insertEmptyRow(worksheetXml: string, rowNumber: number, attributes: string): string {
  const newRow = `<row r="${rowNumber}"${attributes}/>`
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
  throw new StructuralShiftError('Worksheet has no sheetData element.')
}

interface ColDefinition {
  min: number
  max: number
  /// Attributes other than min/max, in original order.
  attrs: Map<string, string>
}

/// Rewrites the `<cols>` section for a width or visibility change. Existing
/// `<col>` ranges overlapping the span are split so untouched columns keep
/// their exact attributes.
function applyColAttributeOp(xml: string, op: AxisAttributeOp): string {
  const existing: ColDefinition[] = []
  const section = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(xml)
  if (section?.[1]) {
    for (const element of section[1].matchAll(/<col\b([^>]*?)\/>/g)) {
      const attrs = new Map<string, string>()
      for (const attribute of (element[1] ?? '').matchAll(/([\w:]+)="([^"]*)"/g)) {
        if (attribute[1] && attribute[2] !== undefined) attrs.set(attribute[1], attribute[2])
      }
      const min = Number(attrs.get('min'))
      const max = Number(attrs.get('max'))
      if (!Number.isInteger(min) || !Number.isInteger(max)) continue
      attrs.delete('min')
      attrs.delete('max')
      existing.push({ min: min - 1, max: max - 1, attrs })
    }
  }

  const patch = (attrs: Map<string, string>): Map<string, string> => {
    const next = new Map(attrs)
    if ('size' in op) {
      next.delete('width')
      next.delete('customWidth')
      next.delete('bestFit')
      if (op.size !== null) {
        next.set('width', String(op.size))
        next.set('customWidth', '1')
      }
    } else if ('level' in op) {
      next.delete('outlineLevel')
      if (op.level > 0) next.set('outlineLevel', String(op.level))
      if (op.collapsed !== undefined) {
        next.delete('collapsed')
        if (op.collapsed) next.set('collapsed', '1')
      }
    } else {
      next.delete('hidden')
      if (op.hidden) next.set('hidden', '1')
    }
    return next
  }

  const result: ColDefinition[] = []
  for (const col of existing) {
    if (col.max < op.start || col.min > op.end) {
      result.push(col)
      continue
    }
    if (col.min < op.start)
      result.push({ min: col.min, max: op.start - 1, attrs: new Map(col.attrs) })
    result.push({
      min: Math.max(col.min, op.start),
      max: Math.min(col.max, op.end),
      attrs: patch(col.attrs),
    })
    if (col.max > op.end) result.push({ min: op.end + 1, max: col.max, attrs: new Map(col.attrs) })
  }
  // Span parts no existing <col> covered get an element carrying only the
  // change (unless the change is a reset back to the default).
  const setsSomething =
    'size' in op
      ? op.size !== null
      : 'level' in op
        ? op.level > 0 || op.collapsed === true
        : op.hidden
  if (setsSomething) {
    const covered = result
      .filter((col) => col.max >= op.start && col.min <= op.end)
      .sort((left, right) => left.min - right.min)
    let cursor = op.start
    const gaps: Array<{ min: number; max: number }> = []
    for (const col of covered) {
      if (col.min > cursor) gaps.push({ min: cursor, max: col.min - 1 })
      cursor = Math.max(cursor, col.max + 1)
    }
    if (cursor <= op.end) gaps.push({ min: cursor, max: op.end })
    for (const gap of gaps) {
      result.push({ min: gap.min, max: gap.max, attrs: patch(new Map()) })
    }
  }

  const kept = result
    .filter((col) => col.attrs.size > 0)
    .sort((left, right) => left.min - right.min)
  const serialized = kept
    .map((col) => {
      const attributes = [...col.attrs].map(([name, value]) => ` ${name}="${value}"`).join('')
      return `<col min="${col.min + 1}" max="${col.max + 1}"${attributes}/>`
    })
    .join('')
  const newSection = kept.length === 0 ? '' : `<cols>${serialized}</cols>`
  if (section) {
    return xml.replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/, () => newSection)
  }
  if (newSection === '') return xml
  // Schema order places cols immediately before sheetData.
  const sheetData = /<sheetData\s*\/>|<sheetData\b[^>]*>/.exec(xml)
  if (!sheetData) throw new StructuralShiftError('Worksheet has no sheetData element.')
  return xml.slice(0, sheetData.index) + newSection + xml.slice(sheetData.index)
}

function refreshMergeCount(xml: string): string {
  return xml.replace(/<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/, (full, inner: string) => {
    const count = (inner.match(/<mergeCell\b/g) ?? []).length
    if (count === 0) return ''
    return full.replace(
      /(<mergeCells\b[^>]*\bcount=")[0-9]+(")/,
      (_m, prefix: string, suffix: string) => `${prefix}${count}${suffix}`,
    )
  })
}

/// Rewrites references to the edited sheet inside another sheet's formulas
/// (only tokens qualified with the edited sheet's name shift).
export function shiftCrossSheetFormulas(
  otherWorksheetXml: string,
  editedSheetName: string,
  ops: readonly StructuralOp[],
): string {
  let xml = otherWorksheetXml
  for (const op of rowColumnOps(ops)) {
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    // The attribute part must not end with '/': a self-closing shared
    // formula (`<f t="shared" si="0"/>`) is not an opening tag, and matching
    // it would swallow real XML up to the next `</f>` as a "formula body".
    xml = xml.replace(
      /<f\b([^>]*[^/>])?>([\s\S]*?)<\/f>/g,
      (_full, attributes: string | undefined, body: string) =>
        `<f${attributes ?? ''}>${escapeXmlText(
          shiftFormulaText(decodeEntities(body), editedSheetName, shift, axis, true),
        )}</f>`,
    )
  }
  return xml
}

/// Rewrites qualified references in workbook definedNames the same way.
export function shiftDefinedNames(
  workbookXml: string,
  editedSheetName: string,
  ops: readonly StructuralOp[],
): string {
  let xml = workbookXml
  for (const op of rowColumnOps(ops)) {
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    xml = xml.replace(
      /(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g,
      (_full, open: string, body: string, close: string) =>
        `${open}${escapeXmlText(
          shiftFormulaText(decodeEntities(body), editedSheetName, shift, axis, true),
        )}${close}`,
    )
  }
  return xml
}

/// Rewrites chart series references (`<c:f>` elements) to the edited sheet.
export function shiftChartReferences(
  chartXml: string,
  editedSheetName: string,
  ops: readonly StructuralOp[],
): string {
  let xml = chartXml
  for (const op of rowColumnOps(ops)) {
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    xml = xml.replace(
      /(<c:f>)([\s\S]*?)(<\/c:f>)/g,
      (_full, open: string, body: string, close: string) =>
        `${open}${escapeXmlText(
          shiftFormulaText(decodeEntities(body), editedSheetName, shift, axis, true),
        )}${close}`,
    )
  }
  return xml
}

function rowColumnOps(ops: readonly StructuralOp[]): RowColumnOp[] {
  return ops.filter((op): op is RowColumnOp => 'index' in op)
}

/// Shifts twoCellAnchor/oneCellAnchor from/to markers in a drawing part along
/// the same op stream. Markers inside a deleted span clamp to its start (with
/// a zeroed offset) — objects are never dropped, only compressed.
/// absoluteAnchor is EMU-positioned and unaffected.
export function shiftDrawingAnchors(drawingXml: string, ops: readonly StructuralOp[]): string {
  let xml = drawingXml
  for (const op of rowColumnOps(ops)) {
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    const tag = axis === 'row' ? 'row' : 'col'
    if (shift.swap) {
      xml = swapDrawingAnchors(xml, shift, tag)
      continue
    }
    xml = xml.replace(
      /<((?:\w+:)?)(from|to)>([\s\S]*?)<\/\1\2>/g,
      (_full, ns: string, kind: string, inner: string) => {
        let clamped = false
        let patched = inner.replace(
          new RegExp(`(<(?:\\w+:)?${tag}>)([0-9]+)(</(?:\\w+:)?${tag}>)`),
          (_m, open: string, value: string, close: string) => {
            const moved = moveAnchorMark(Number(value), shift)
            clamped = moved.clamped
            return `${open}${moved.position}${close}`
          },
        )
        if (clamped) {
          patched = patched.replace(
            new RegExp(`(<(?:\\w+:)?${tag}Off>)-?[0-9]+(</(?:\\w+:)?${tag}Off>)`),
            (_m, open: string, close: string) => `${open}0${close}`,
          )
        }
        return `<${ns}${kind}>${patched}</${ns}${kind}>`
      },
    )
  }
  return xml
}

/// Swap-shift pass over a drawing part: from/to marks are judged as a PAIR
/// (independent remapping could tear an anchor whose ends straddle a block
/// boundary). moveRange supplies the three-way semantics: anchors outside or
/// spanning the swapped blocks stay put, anchors inside one block move with
/// it, partial contact fails closed.
function swapDrawingAnchors(xml: string, shift: BlockSwap, tag: string): string {
  const markPattern = (kind: string): RegExp =>
    new RegExp(`(<(?:\\w+:)?${kind}>[\\s\\S]*?<(?:\\w+:)?${tag}>)([0-9]+)(</(?:\\w+:)?${tag}>)`)
  const result = xml.replace(/<((?:\w+:)?)twoCellAnchor\b[\s\S]*?<\/\1twoCellAnchor>/g, (block) => {
    const from = markPattern('from').exec(block)
    const to = markPattern('to').exec(block)
    if (!from?.[2] || !to?.[2]) return block
    const moved = moveRange(Number(from[2]), Number(to[2]), shift)
    if (!moved || (moved.start === Number(from[2]) && moved.end === Number(to[2]))) return block
    return block
      .replace(
        markPattern('from'),
        (_m, open: string, _v, close: string) => `${open}${moved.start}${close}`,
      )
      .replace(
        markPattern('to'),
        (_m, open: string, _v, close: string) => `${open}${moved.end}${close}`,
      )
  })
  // oneCellAnchor is extent-based; its single mark remaps freely.
  return result.replace(/<((?:\w+:)?)oneCellAnchor\b[\s\S]*?<\/\1oneCellAnchor>/g, (block) => {
    const from = markPattern('from').exec(block)
    if (!from?.[2]) return block
    const moved = movePosition(Number(from[2]), shift)
    if (moved === null || moved === Number(from[2])) return block
    return block.replace(
      markPattern('from'),
      (_m, open: string, _v, close: string) => `${open}${moved}${close}`,
    )
  })
}

function moveAnchorMark(
  position: number,
  shift: LinearShift,
): { position: number; clamped: boolean } {
  if (shift.deleted) {
    if (position > shift.deleted.end) return { position: position + shift.delta, clamped: false }
    if (position >= shift.deleted.start) return { position: shift.deleted.start, clamped: true }
    return { position, clamped: false }
  }
  return {
    position: position >= shift.boundary ? position + shift.delta : position,
    clamped: false,
  }
}

/// Shifts a table part's ref/autoFilter/sortState ranges along the op stream.
/// Fail-closed cases (the tableColumns list or table anatomy would have to
/// change): deleting header/totals rows or the whole table, deleting all data
/// rows, inserting/deleting columns inside the table, merging over it.
export function shiftTablePart(tableXml: string, ops: readonly StructuralOp[]): string {
  let xml = tableXml
  for (const op of ops) {
    if ('start' in op) continue
    const table = parseTablePart(xml)
    if ('range' in op) {
      if (op.kind === 'merge-cells' && areasOverlap(op.range, table)) {
        throw new StructuralShiftError(`Merging cells over table "${table.name}" is not supported.`)
      }
      continue
    }
    const axis: Axis = axisOf(op)
    const shift = toShift(op)
    if (axis === 'row') assertTableRowShiftSupported(table, shift)
    else assertTableColumnShiftSupported(table, shift)
    xml = xml.replace(
      /(<(?:table|autoFilter|sortState|sortCondition)\b[^>]*?\bref=")([^"]+)(")/g,
      (full, prefix: string, ref: string, suffix: string) => {
        const moved = moveRefRange(ref, shift, axis)
        return moved === null ? full : `${prefix}${moved}${suffix}`
      },
    )
  }
  return xml
}

interface TablePartArea extends CellArea {
  readonly name: string
  readonly headerRows: number
  readonly totalsRows: number
}

function parseTablePart(tableXml: string): TablePartArea {
  const open = /<table\b[^>]*>/.exec(tableXml)?.[0]
  const ref = open === undefined ? undefined : /\bref="([^"]+)"/.exec(open)?.[1]
  const [startRef, endRef] = (ref ?? '').split(':')
  const start = parseA1(startRef ?? '')
  const end = parseA1(endRef ?? startRef ?? '')
  if (!open || !start || !end) {
    throw new StructuralShiftError('An Excel table part has no readable ref — aborted.')
  }
  return {
    name: /\bname="([^"]*)"/.exec(open)?.[1] ?? 'Table',
    startRow: start.row,
    endRow: end.row,
    startColumn: start.column,
    endColumn: end.column,
    headerRows: Number(/\bheaderRowCount="([0-9]+)"/.exec(open)?.[1] ?? '1'),
    totalsRows: Number(/\btotalsRowCount="([0-9]+)"/.exec(open)?.[1] ?? '0'),
  }
}

function spansOverlap(aStart: number, aEnd: number, bStart: number, bEnd: number): boolean {
  return aStart <= bEnd && bStart <= aEnd
}

function areasOverlap(a: CellArea, b: CellArea): boolean {
  return (
    spansOverlap(a.startRow, a.endRow, b.startRow, b.endRow) &&
    spansOverlap(a.startColumn, a.endColumn, b.startColumn, b.endColumn)
  )
}

function assertTableRowShiftSupported(table: TablePartArea, shift: Shift): void {
  if (shift.swap) {
    assertTableRowMoveSupported(table, shift.swap)
    return
  }
  if (!shift.deleted) return
  const { start, end } = shift.deleted
  if (
    table.headerRows > 0 &&
    spansOverlap(start, end, table.startRow, table.startRow + table.headerRows - 1)
  ) {
    throw new StructuralShiftError(
      `Deleting the header row of table "${table.name}" is not supported.`,
    )
  }
  if (
    table.totalsRows > 0 &&
    spansOverlap(start, end, table.endRow - table.totalsRows + 1, table.endRow)
  ) {
    throw new StructuralShiftError(
      `Deleting the totals row of table "${table.name}" is not supported.`,
    )
  }
  const moved = moveRange(table.startRow, table.endRow, shift)
  const dataRows =
    moved === null ? 0 : moved.end - moved.start + 1 - table.headerRows - table.totalsRows
  if (dataRows < 1) {
    throw new StructuralShiftError(
      `Deleting these rows would leave table "${table.name}" without data rows — not supported.`,
    )
  }
}

/// Whole-table relocation (table fully inside one swapped block) and data-row
/// reorders inside the table are fine; a move that would relocate the header
/// or totals row relative to the data fails closed. Boundary tears are left
/// to moveRange on the table ref.
function assertTableRowMoveSupported(table: TablePartArea, swap: BlockSwap['swap']): void {
  const within = (span: Span): boolean => table.startRow >= span.start && table.endRow <= span.end
  if (within(swap.first) || within(swap.second)) return
  const overlapsEnvelope = (start: number, end: number): boolean =>
    spansOverlap(start, end, swap.first.start, swap.second.end)
  if (!overlapsEnvelope(table.startRow, table.endRow)) return
  if (
    table.headerRows > 0 &&
    overlapsEnvelope(table.startRow, table.startRow + table.headerRows - 1)
  ) {
    throw new StructuralShiftError(
      `Moving the header row of table "${table.name}" is not supported.`,
    )
  }
  if (table.totalsRows > 0 && overlapsEnvelope(table.endRow - table.totalsRows + 1, table.endRow)) {
    throw new StructuralShiftError(
      `Moving the totals row of table "${table.name}" is not supported.`,
    )
  }
}

function assertTableColumnShiftSupported(table: TablePartArea, shift: Shift): void {
  if (shift.swap) return
  if (shift.deleted) {
    if (spansOverlap(shift.deleted.start, shift.deleted.end, table.startColumn, table.endColumn)) {
      throw new StructuralShiftError(
        `Deleting columns of table "${table.name}" is not supported — delete the table first.`,
      )
    }
    return
  }
  if (shift.boundary > table.startColumn && shift.boundary <= table.endColumn) {
    throw new StructuralShiftError(
      `Inserting columns inside table "${table.name}" is not supported.`,
    )
  }
}

type Axis = 'row' | 'column'

interface Span {
  readonly start: number
  readonly end: number
}

interface LinearShift {
  readonly swap?: undefined
  /// 0-based first affected index on the op's axis.
  readonly boundary: number
  /// positive for insert, negative for delete
  readonly delta: number
  /// inclusive deleted index range in the pre-op space, or null for inserts
  readonly deleted: { start: number; end: number } | null
}

/// A whole-row move expressed as two adjacent index blocks swapping places —
/// a bijection over the axis, so positions never map to null.
interface BlockSwap {
  readonly swap: { readonly first: Span; readonly second: Span }
}

type Shift = LinearShift | BlockSwap

function axisOf(op: RowColumnOp): Axis {
  return op.kind === 'insert-cols' || op.kind === 'remove-cols' ? 'column' : 'row'
}

function toShift(op: RowColumnOp): Shift {
  if (op.kind === 'move-rows') {
    if (op.before >= op.index && op.before <= op.index + op.count) {
      throw new StructuralShiftError('A row move has a target inside the moved block — aborted.')
    }
    return op.before > op.index
      ? {
          swap: {
            first: { start: op.index, end: op.index + op.count - 1 },
            second: { start: op.index + op.count, end: op.before - 1 },
          },
        }
      : {
          swap: {
            first: { start: op.before, end: op.index - 1 },
            second: { start: op.index, end: op.index + op.count - 1 },
          },
        }
  }
  if (op.kind === 'insert-rows' || op.kind === 'insert-cols') {
    return { boundary: op.index, delta: op.count, deleted: null }
  }
  return {
    boundary: op.index,
    delta: -op.count,
    deleted: { start: op.index, end: op.index + op.count - 1 },
  }
}

/// Shifts a single 0-based position; null means it was deleted.
function movePosition(position: number, shift: Shift): number | null {
  if (shift.swap) {
    const { first, second } = shift.swap
    if (position >= first.start && position <= first.end) {
      return position + (second.end - second.start + 1)
    }
    if (position >= second.start && position <= second.end) {
      return position - (first.end - first.start + 1)
    }
    return position
  }
  if (shift.deleted) {
    if (position >= shift.deleted.start && position <= shift.deleted.end) return null
    return position > shift.deleted.end ? position + shift.delta : position
  }
  return position >= shift.boundary ? position + shift.delta : position
}

/// Shifts an inclusive 0-based range. Deleting part of it shrinks it (Excel
/// semantics for merges, scoped ranges, and SUM-style range references);
/// null only when the whole range is deleted. Inserting inside expands it.
/// For swaps the mapping is three-way: ranges outside or containing the
/// swapped blocks stay put, ranges inside one block move with it, and
/// anything partially overlapping fails closed — there is no faithful
/// single-range image of a torn range.
function moveRange(
  start: number,
  end: number,
  shift: Shift,
): { start: number; end: number } | null {
  if (shift.swap) {
    const { first, second } = shift.swap
    if (end < first.start || start > second.end) return { start, end }
    if (start <= first.start && end >= second.end) return { start, end }
    if (start >= first.start && end <= first.end) {
      const delta = second.end - second.start + 1
      return { start: start + delta, end: end + delta }
    }
    if (start >= second.start && end <= second.end) {
      const delta = first.end - first.start + 1
      return { start: start - delta, end: end - delta }
    }
    throw new StructuralShiftError(
      'A range partially overlaps the moved rows — the move cannot be saved.',
    )
  }
  if (shift.deleted) {
    const { start: ds, end: de } = shift.deleted
    if (start >= ds && end <= de) return null
    const newStart = start > de ? start + shift.delta : start >= ds ? ds : start
    const newEnd = end > de ? end + shift.delta : end >= ds ? ds - 1 : end
    return newEnd < newStart ? null : { start: newStart, end: newEnd }
  }
  return {
    start: start >= shift.boundary ? start + shift.delta : start,
    end: end >= shift.boundary ? end + shift.delta : end,
  }
}

function transformSheetRows(xml: string, shift: Shift): string {
  const renumbered = xml.replace(
    /<row\b[^>]*?\br="([0-9]+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g,
    (full, rowNumber: string) => {
      const moved = movePosition(Number(rowNumber) - 1, shift)
      if (moved === null) return ''
      if (moved === Number(rowNumber) - 1) return full
      const patched = full.replace(
        /(<row\b[^>]*?\br=")[0-9]+(")/,
        (_m, prefix: string, suffix: string) => `${prefix}${moved + 1}${suffix}`,
      )
      // Cell addresses inside the row carry the row number too.
      return patched.replace(
        /(<c\b[^>]*?\br="[A-Z]{1,3})[0-9]+(")/g,
        (_m, prefix: string, suffix: string) => `${prefix}${moved + 1}${suffix}`,
      )
    },
  )
  // Inserts and deletes renumber monotonically, so document order survives;
  // a swap does not — Excel requires sheetData rows in ascending order.
  return shift.swap ? sortSheetDataRows(renumbered) : renumbered
}

function sortSheetDataRows(xml: string): string {
  const section = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml)
  if (!section?.[1]) return xml
  const rows: { index: number; text: string }[] = []
  const leftover = section[1].replace(
    /<row\b[^>]*?\br="([0-9]+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g,
    (full, rowNumber: string) => {
      rows.push({ index: Number(rowNumber), text: full })
      return ''
    },
  )
  if (leftover.trim() !== '') {
    throw new StructuralShiftError('sheetData holds content other than rows — move aborted.')
  }
  rows.sort((a, b) => a.index - b.index)
  const body = rows.map((row) => row.text).join('')
  return (
    xml.slice(0, section.index) +
    `${section[0].slice(0, section[0].indexOf('>') + 1)}${body}</sheetData>` +
    xml.slice(section.index + section[0].length)
  )
}

function transformSheetColumns(xml: string, shift: Shift): string {
  // Row span hints go stale after a column shift; Excel rebuilds them.
  let result = xml.replace(/(<row\b[^>]*?)\s+spans="[^"]*"/g, (_m, prefix: string) => prefix)
  result = result.replace(
    /<c\b[^>]*?\br="([A-Z]{1,3})[0-9]+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,
    (full, letters: string) => {
      const moved = movePosition(lettersToColumn(letters), shift)
      if (moved === null) return ''
      if (moved === lettersToColumn(letters)) return full
      return full.replace(
        /(<c\b[^>]*?\br=")[A-Z]{1,3}([0-9]+")/,
        (_m, prefix: string, suffix: string) => `${prefix}${columnToLetters(moved)}${suffix}`,
      )
    },
  )
  return result
}

function transformColDefinitions(xml: string, shift: Shift): string {
  return xml.replace(
    /<col\b([^>]*?)\bmin="([0-9]+)"([^>]*?)\bmax="([0-9]+)"([^>]*?)\/>/g,
    (_full, b1: string, min: string, b2: string, max: string, b3: string) => {
      const moved = moveRange(Number(min) - 1, Number(max) - 1, shift)
      if (moved === null) return ''
      return `<col${b1}min="${moved.start + 1}"${b2}max="${moved.end + 1}"${b3}/>`
    },
  )
}

/// Rewrites `<f>` bodies plus shared/array formula `ref` attributes, and the
/// `<formula>`/`<formula1>`/`<formula2>` bodies of conditional formatting and
/// data validation rules.
function transformFormulas(xml: string, sheetName: string, shift: Shift, axis: Axis): string {
  // Self-closing `<f .../>` must not match — see shiftCrossSheetFormulas.
  let result = xml.replace(
    /<f\b([^>]*[^/>])?>([\s\S]*?)<\/f>/g,
    (_full, rawAttributes: string | undefined, body: string) => {
      const attributes = rawAttributes ?? ''
      const rewritten = escapeXmlText(
        shiftFormulaText(decodeEntities(body), sheetName, shift, axis),
      )
      const newAttributes = attributes.replace(
        /(\bref=")([^"]*)(")/,
        (_m, prefix: string, ref: string, suffix: string) => {
          // A shared/array anchor spanning the swapped blocks would have its
          // si expansion reordered underneath it — only wholesale moves of
          // the anchor (or no contact at all) are safe.
          if (shift.swap && axis === 'row') assertSwapKeepsAnchorIntact(ref, shift.swap)
          const moved = moveRefRange(ref, shift, axis)
          if (moved === null) {
            throw new StructuralShiftError(
              'Deletion would remove a shared formula anchor — aborted.',
            )
          }
          return `${prefix}${moved}${suffix}`
        },
      )
      return `<f${newAttributes}>${rewritten}</f>`
    },
  )
  result = result.replace(
    /<(formula[12]?)>([\s\S]*?)<\/\1>/g,
    (_full, tag: string, body: string) =>
      `<${tag}>${escapeXmlText(shiftFormulaText(decodeEntities(body), sheetName, shift, axis))}</${tag}>`,
  )
  return result
}

/// Shifts merges and range-scoped feature elements. Elements whose every
/// range is deleted are removed outright.
function transformRangedFeatures(xml: string, shift: Shift, axis: Axis): string {
  let result = xml.replace(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/>/g, (full, ref: string) => {
    const moved = moveRefRange(ref, shift, axis)
    return moved === null ? '' : full.replace(/ref="[^"]+"/, () => `ref="${moved}"`)
  })
  result = result.replace(
    /<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>|<mergeCells\b[^>]*\/>/g,
    (full, inner: string | undefined) => {
      const count = ((inner ?? '').match(/<mergeCell\b/g) ?? []).length
      if (count === 0) return ''
      return full.replace(
        /(<mergeCells\b[^>]*\bcount=")[0-9]+(")/,
        (_m, prefix: string, suffix: string) => `${prefix}${count}${suffix}`,
      )
    },
  )
  // dimension / autoFilter shift in place (kept even when degenerate).
  result = result.replace(
    /(<(?:dimension|autoFilter)\b[^>]*?\bref=")([^"]+)(")/g,
    (full, prefix: string, ref: string, suffix: string) => {
      const moved = moveRefRange(ref, shift, axis)
      return moved === null ? full : `${prefix}${moved}${suffix}`
    },
  )
  for (const tag of ['hyperlink', 'dataValidation', 'conditionalFormatting']) {
    const attribute = tag === 'hyperlink' ? 'ref' : 'sqref'
    result = result.replace(
      new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>|<${tag}\\b[^>]*/>`, 'g'),
      (element) => {
        const refMatch = new RegExp(`\\b${attribute}="([^"]+)"`).exec(element)
        if (!refMatch?.[1]) return element
        const moved = refMatch[1]
          .split(' ')
          .map((ref) => moveRefRange(ref, shift, axis))
          .filter((ref): ref is string => ref !== null)
        if (moved.length === 0) return ''
        return element.replace(
          new RegExp(`\\b${attribute}="[^"]+"`),
          () => `${attribute}="${moved.join(' ')}"`,
        )
      },
    )
  }
  result = result.replace(
    /(<dataValidations\b[^>]*\bcount=")[0-9]+("[^>]*>)([\s\S]*?)(<\/dataValidations>)/g,
    (_full, prefix: string, mid: string, inner: string, close: string) => {
      const count = (inner.match(/<dataValidation\b/g) ?? []).length
      return count === 0 ? '' : `${prefix}${count}${mid}${inner}${close}`
    },
  )
  return result
}

function assertSwapKeepsAnchorIntact(ref: string, swap: BlockSwap['swap']): void {
  const parts = ref.split(':')
  const start = parseA1(parts[0] ?? '')
  const end = parseA1(parts[1] ?? parts[0] ?? '')
  if (!start || !end) return
  const outside = end.row < swap.first.start || start.row > swap.second.end
  const insideFirst = start.row >= swap.first.start && end.row <= swap.first.end
  const insideSecond = start.row >= swap.second.start && end.row <= swap.second.end
  if (outside || insideFirst || insideSecond) return
  throw new StructuralShiftError(
    'A shared formula anchor overlaps the moved rows — the move cannot be saved.',
  )
}

function moveRefRange(ref: string, shift: Shift, axis: Axis): string | null {
  const parts = ref.split(':')
  const start = parseA1(parts[0] ?? '')
  if (!start) return ref
  if (parts.length === 1) {
    const moved = movePosition(axis === 'row' ? start.row : start.column, shift)
    if (moved === null) return null
    return axis === 'row'
      ? `${columnToLetters(start.column)}${moved + 1}`
      : `${columnToLetters(moved)}${start.row + 1}`
  }
  const end = parseA1(parts[1] ?? '')
  if (!end) return ref
  const moved =
    axis === 'row'
      ? moveRange(start.row, end.row, shift)
      : moveRange(start.column, end.column, shift)
  if (moved === null) return null
  return axis === 'row'
    ? `${columnToLetters(start.column)}${moved.start + 1}:${columnToLetters(end.column)}${moved.end + 1}`
    : `${columnToLetters(moved.start)}${start.row + 1}:${columnToLetters(moved.end)}${end.row + 1}`
}

// A1-style token, optionally sheet-qualified: cell, cell range, whole-column
// range, or whole-row range. The leading guard keeps defined-name characters
// from being misread as references.
export const FORMULA_REFERENCE_PATTERN = new RegExp(
  "(^|[^A-Za-z0-9_.$'!:])" +
    "(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?" +
    '(\\$?[A-Z]{1,3}\\$?[0-9]+(?::\\$?[A-Z]{1,3}\\$?[0-9]+)?' +
    '|\\$?[A-Z]{1,3}:\\$?[A-Z]{1,3}' +
    '|\\$?[0-9]+:\\$?[0-9]+)' +
    '(?![0-9A-Za-z_(])',
  'g',
)

/// Shifts A1 references in a formula, preserving `$` markers and other
/// sheets' qualified references. String literals are left untouched. A
/// reference fully inside a deleted range aborts (Excel would emit #REF!).
/// With qualifiedOnly, only references explicitly qualified with sheetName
/// shift — the mode for rewriting OTHER sheets' formulas.
export function shiftFormulaText(
  formula: string,
  sheetName: string,
  shift: Shift,
  axis: Axis,
  qualifiedOnly = false,
): string {
  // Formula string literals use "" escaping, so splitting on `"` leaves
  // literal content in the odd-indexed segments.
  return formula
    .split('"')
    .map((segment, index) =>
      index % 2 === 1
        ? segment
        : shiftFormulaSegment(segment, sheetName, shift, axis, qualifiedOnly),
    )
    .join('"')
}

function shiftFormulaSegment(
  segment: string,
  sheetName: string,
  shift: Shift,
  axis: Axis,
  qualifiedOnly: boolean,
): string {
  return segment.replace(
    FORMULA_REFERENCE_PATTERN,
    (full, lead: string, qualifier: string | undefined, token: string) => {
      if (qualifier === undefined && qualifiedOnly) return full
      if (qualifier !== undefined && !qualifierMatches(qualifier, sheetName)) return full
      const shifted = shiftReferenceToken(token, shift, axis)
      if (shifted === null) {
        throw new StructuralShiftError(
          `A formula references the deleted range (${token}) — deletion aborted.`,
        )
      }
      return `${lead}${qualifier === undefined ? '' : `${qualifier}!`}${shifted}`
    },
  )
}

function shiftReferenceToken(token: string, shift: Shift, axis: Axis): string | null {
  const wholeColumn = /^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/.exec(token)
  if (wholeColumn) {
    if (axis === 'row') return token
    const moved = moveRange(
      lettersToColumn(wholeColumn[2] ?? 'A'),
      lettersToColumn(wholeColumn[4] ?? 'A'),
      shift,
    )
    if (moved === null) return null
    return `${wholeColumn[1]}${columnToLetters(moved.start)}:${wholeColumn[3]}${columnToLetters(moved.end)}`
  }
  const wholeRow = /^(\$?)([0-9]+):(\$?)([0-9]+)$/.exec(token)
  if (wholeRow) {
    if (axis === 'column') return token
    const moved = moveRange(Number(wholeRow[2]) - 1, Number(wholeRow[4]) - 1, shift)
    if (moved === null) return null
    return `${wholeRow[1]}${moved.start + 1}:${wholeRow[3]}${moved.end + 1}`
  }
  const parts = token.split(':')
  const cells = parts.map((part) => /^(\$?)([A-Z]{1,3})(\$?)([0-9]+)$/.exec(part))
  if (cells.some((cell) => cell === null)) return token
  const parsed = cells.map((cell) => ({
    colDollar: cell?.[1] ?? '',
    column: lettersToColumn(cell?.[2] ?? 'A'),
    rowDollar: cell?.[3] ?? '',
    row: Number(cell?.[4]) - 1,
  }))
  const first = parsed[0]
  if (!first) return token
  if (parsed.length === 1) {
    const moved = movePosition(axis === 'row' ? first.row : first.column, shift)
    if (moved === null) return null
    const row = axis === 'row' ? moved : first.row
    const column = axis === 'column' ? moved : first.column
    return `${first.colDollar}${columnToLetters(column)}${first.rowDollar}${row + 1}`
  }
  const second = parsed[1]
  if (!second) return token
  const moved =
    axis === 'row'
      ? moveRange(first.row, second.row, shift)
      : moveRange(first.column, second.column, shift)
  if (moved === null) return null
  const startRow = axis === 'row' ? moved.start : first.row
  const endRow = axis === 'row' ? moved.end : second.row
  const startColumn = axis === 'column' ? moved.start : first.column
  const endColumn = axis === 'column' ? moved.end : second.column
  return (
    `${first.colDollar}${columnToLetters(startColumn)}${first.rowDollar}${startRow + 1}` +
    `:${second.colDollar}${columnToLetters(endColumn)}${second.rowDollar}${endRow + 1}`
  )
}

export function qualifierMatches(qualifier: string, sheetName: string): boolean {
  const unquoted = qualifier.startsWith("'")
    ? qualifier.slice(1, -1).replaceAll("''", "'")
    : qualifier
  return unquoted === sheetName
}

function parseA1(ref: string): { row: number; column: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(ref)
  if (!match?.[1] || !match[2]) return null
  return { row: Number(match[2]) - 1, column: lettersToColumn(match[1]) }
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
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

function decodeEntities(input: string): string {
  return input
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&gt;', '>')
    .replaceAll('&lt;', '<')
    .replaceAll('&amp;', '&')
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}
