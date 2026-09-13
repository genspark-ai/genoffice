import type {
  PivotDataField,
  PivotDefinition,
  PivotFieldItem,
  PivotLayoutLine,
  PivotSharedItem,
} from '../gateway/xlsx-pivot'
import { allowedByValueFilter, matchesLabelFilter } from './pivot-filters'
import { evaluatePivotFormula, parsePivotFormula, type PivotFormulaAst } from './pivot-formula'
import { groupLabel } from './pivot-grouping'

/// Recomputes a pivot table's DATA AREA from fresh source values, keeping the
/// file's recorded layout (rowItems/colItems). Layout drift — new or renamed
/// categories in the source — fails closed: growing the layout is a separate
/// concern (M2), a silently truncated refresh would lie.

export class PivotRefreshError extends Error {}

export interface PivotRecomputeResult {
  /// Data-area grid (rows × columns), positioned at (firstDataRow,
  /// firstDataCol) inside outputRef. null = leave the cell blank.
  readonly data: readonly (readonly (number | null)[])[]
}

type SourceCell = string | number | boolean | null | undefined

/// Case-insensitive string grouping for pivot buckets.
function groupKey(value: PivotSharedItem | SourceCell): string {
  if (value === null || value === undefined || value === '') return '\u0000blank'
  if (typeof value === 'number') return `n:${value}`
  if (typeof value === 'boolean') return `b:${value}`
  const numeric = Number(value)
  if (String(value).trim() !== '' && Number.isFinite(numeric)) return `n:${numeric}`
  return `s:${String(value).trim().toLowerCase()}`
}

interface AxisPredicate {
  readonly field: number
  readonly key: string
}

/// Grouping key of a source cell on a dimension field: grouped fields
/// (date/numeric ranges) are first bucketed into labels, then keyed with pivot
/// semantics — the cache's sharedItems are the group labels, so both sides agree.
function sourceKey(definition: PivotDefinition, field: number, value: SourceCell): string {
  const grouping = definition.fields[field]?.grouping
  return groupKey(grouping ? groupLabel(grouping, value) : value)
}

/// Fixed members of a layout line as source-value predicates. The data-field
/// axis position (-2) selects which value column renders, not a member.
function linePredicates(
  definition: PivotDefinition,
  axisFields: readonly number[],
  line: PivotLayoutLine,
): { predicates: AxisPredicate[]; dataField: number } {
  const predicates: AxisPredicate[] = []
  let dataField = line.dataField
  const levels = line.t === 'data' ? axisFields.length : line.depth
  for (let level = 0; level < levels; level += 1) {
    const field = axisFields[level]
    const member = line.members[level]
    if (field === undefined || member === null || member === undefined) continue
    if (field === -2) {
      dataField = member
      continue
    }
    const item = definition.fieldItems[field]?.[member]
    if (!item || item.x === null) {
      throw new PivotRefreshError(`Item ${member} of field ${field} cannot be resolved.`)
    }
    const shared = definition.fields[field]?.sharedItems[item.x]
    if (shared === undefined) {
      throw new PivotRefreshError(`Shared item ${item.x} of field ${field} is missing.`)
    }
    predicates.push({ field, key: groupKey(shared) })
  }
  return { predicates, dataField }
}

function aggregate(subtotal: string, values: readonly SourceCell[]): number | null {
  if (subtotal === 'count') {
    return values.filter((value) => value !== null && value !== undefined && value !== '').length
  }
  const numbers = values
    .map((value) => (typeof value === 'number' ? value : Number(String(value ?? '').trim())))
    .filter((value, index) => {
      const raw = values[index]
      return raw !== null && raw !== undefined && raw !== '' && Number.isFinite(value)
    })
  switch (subtotal) {
    case 'countNums':
      return numbers.length
    case 'sum':
      return numbers.reduce((total, value) => total + value, 0)
    case 'average':
      return numbers.length === 0 ? null : numbers.reduce((t, v) => t + v, 0) / numbers.length
    case 'max':
      return numbers.length === 0 ? null : Math.max(...numbers)
    case 'min':
      return numbers.length === 0 ? null : Math.min(...numbers)
    case 'product':
      return numbers.length === 0 ? null : numbers.reduce((t, v) => t * v, 1)
    default:
      throw new PivotRefreshError(`Aggregation "${subtotal}" does not support recompute.`)
  }
}

/// Consistency check between source data and the pivot cache (header width and
/// names): a moved source or renamed header always fails closed. Shared by
/// recompute and layout growth.
function assertSourceMatchesCache(
  definition: PivotDefinition,
  sourceValues: readonly (readonly SourceCell[])[],
): void {
  if (definition.unsupported.length > 0) {
    throw new PivotRefreshError(
      `This pivot table does not support recompute: ${definition.unsupported.join('; ')}`,
    )
  }
  // Calculated fields take no source-data column (the parser guarantees they
  // come after source fields), so validation covers only real source fields.
  const sourceFieldCount = definition.fields.filter((field) => field.formula === undefined).length
  const header = sourceValues[0]
  if (!header || header.length < sourceFieldCount) {
    throw new PivotRefreshError(
      'The source data region is narrower than the pivot cache field count — the data source may have moved.',
    )
  }
  definition.fields.forEach((field, index) => {
    if (field.formula !== undefined) return
    const label = String(header[index] ?? '').trim()
    if (label.toLowerCase() !== field.name.trim().toLowerCase()) {
      throw new PivotRefreshError(
        `Source data column ${index + 1} header "${label}" does not match cache field "${field.name}" — the data source may have moved.`,
      )
    }
  })
}

/// Cache fields involved in axes and report filters (excluding the data-field axis -2).
function axisFields(definition: PivotDefinition): Set<number> {
  const fields = new Set<number>([
    ...definition.rowFields,
    ...definition.colFields,
    ...definition.pageFields.map((page) => page.field),
  ])
  fields.delete(-2)
  return fields
}

/// sourceValues: the worksheetSource range INCLUDING its header row.
export function recomputePivotData(
  definition: PivotDefinition,
  sourceValues: readonly (readonly SourceCell[])[],
): PivotRecomputeResult {
  assertSourceMatchesCache(definition, sourceValues)
  const rows = sourceValues.slice(1)

  // Layout drift: a new distinct value in an axis/page column cannot be
  // placed into the recorded layout. Growth is a separate step
  // (growPivotDefinition) — recompute itself stays fail-closed; callers grow
  // first, then recompute.
  const axisFieldSet = axisFields(definition)
  for (const field of axisFieldSet) {
    const known = new Set(definition.fields[field]?.sharedItems.map(groupKey) ?? [])
    for (const row of rows) {
      if (!known.has(sourceKey(definition, field, row[field]))) {
        throw new PivotRefreshError(
          `The source data column "${definition.fields[field]?.name}" contains a new item not in the cache ` +
            `("${String(row[field])}") — refresh the layout in Excel before it can be recomputed.`,
        )
      }
    }
  }

  // Hidden items filter out their rows everywhere (including totals).
  const hiddenKeys = new Map<number, Set<string>>()
  for (const field of axisFieldSet) {
    const hidden = new Set<string>()
    for (const item of definition.fieldItems[field] ?? []) {
      if (!item.hidden || item.x === null) continue
      const shared = definition.fields[field]?.sharedItems[item.x]
      if (shared !== undefined) hidden.add(groupKey(shared))
    }
    if (hidden.size > 0) hiddenKeys.set(field, hidden)
  }
  const pagePredicates: AxisPredicate[] = []
  for (const page of definition.pageFields) {
    if (page.item === null) continue
    const item = definition.fieldItems[page.field]?.[page.item]
    const shared =
      item?.x === null || item === undefined
        ? undefined
        : definition.fields[page.field]?.sharedItems[item.x]
    if (shared === undefined) {
      throw new PivotRefreshError('The selected report filter item cannot be resolved.')
    }
    pagePredicates.push({ field: page.field, key: groupKey(shared) })
  }

  const visibleRows = rows.filter((row) => {
    for (const [field, hidden] of hiddenKeys) {
      if (hidden.has(sourceKey(definition, field, row[field]))) return false
    }
    for (const page of pagePredicates) {
      if (sourceKey(definition, page.field, row[page.field]) !== page.key) return false
    }
    return true
  })

  const matchingValues = (predicates: readonly AxisPredicate[], field: number): SourceCell[] =>
    visibleRows
      .filter((row) =>
        predicates.every(
          (predicate) =>
            sourceKey(definition, predicate.field, row[predicate.field]) === predicate.key,
        ),
      )
      .map((row) => row[field])

  // Calculated fields: source fields referenced by the formula resolve by
  // name; each group first takes the SUM of every referenced field, then
  // evaluates the formula. Formulas parse once and are
  // reused (validated at parse time; a failure here is exceptional — fail-closed).
  const fieldIndexByName = new Map<string, number>()
  definition.fields.forEach((field, index) => {
    if (field.formula === undefined) fieldIndexByName.set(field.name, index)
  })
  const formulaAsts: (PivotFormulaAst | null)[] = definition.dataFields.map((dataField) => {
    if (dataField.formula === undefined) return null
    try {
      return parsePivotFormula(dataField.formula, [...fieldIndexByName.keys()])
    } catch (error) {
      throw new PivotRefreshError(
        `The formula of calculated field "${dataField.name}" cannot be parsed: ` +
          `${error instanceof Error ? error.message : String(error)}`,
      )
    }
  })
  // Aggregated value for one (data field × predicate set): plain data fields
  // aggregate by subtotal; calculated fields evaluate the formula over each
  // referenced field's SUM.
  const aggregateCell = (
    dataFieldIndex: number,
    predicates: readonly AxisPredicate[],
  ): number | null => {
    const dataField = definition.dataFields[dataFieldIndex]
    if (!dataField) throw new PivotRefreshError(`Data field ${dataFieldIndex} is missing.`)
    const ast = formulaAsts[dataFieldIndex]
    if (ast) {
      return evaluatePivotFormula(ast, (name) =>
        aggregate('sum', matchingValues(predicates, fieldIndexByName.get(name)!)),
      )
    }
    return aggregate(dataField.subtotal, matchingValues(predicates, dataField.field))
  }

  // Denominator for "show values as" (percentages): re-aggregate by
  // predicates instead of reading grand-total cells — the denominator stays
  // correct even when the layout disables total rows/columns. The same
  // (row line/col line × data field) denominator is reused by many cells, so memoize.
  const baseCache = new Map<string, number | null>()
  const aggregateBase = (
    cacheKey: string,
    dataFieldIndex: number,
    predicates: readonly AxisPredicate[],
  ): number | null => {
    const key = `${cacheKey}:${dataFieldIndex}`
    if (!baseCache.has(key)) {
      baseCache.set(key, aggregateCell(dataFieldIndex, predicates))
    }
    return baseCache.get(key)!
  }

  const data: (number | null)[][] = []
  definition.rowLines.forEach((rowLine, rowIndex) => {
    if (rowLine.t === 'blank') {
      data.push(new Array(definition.colLines.length).fill(null))
      return
    }
    const rowPart = linePredicates(definition, definition.rowFields, rowLine)
    const line: (number | null)[] = []
    definition.colLines.forEach((colLine, colIndex) => {
      if (colLine.t === 'blank') {
        line.push(null)
        return
      }
      const colPart = linePredicates(definition, definition.colFields, colLine)
      const dataFieldIndex = Math.max(rowPart.dataField, colPart.dataField)
      const dataField = definition.dataFields[dataFieldIndex]
      if (!dataField) throw new PivotRefreshError(`Data field ${dataFieldIndex} is missing.`)
      const raw = aggregateCell(dataFieldIndex, [...rowPart.predicates, ...colPart.predicates])
      if (dataField.showDataAs === undefined || raw === null) {
        line.push(raw)
        return
      }
      // Second-pass normalization: % of row = this row's total (row
      // predicates only); % of column = this column's total (column predicates
      // only); % of grand total = all visible rows. Also applies to subtotal
      // and grand-total cells.
      const base =
        dataField.showDataAs === 'percentOfTotal'
          ? aggregateBase('total', dataFieldIndex, [])
          : dataField.showDataAs === 'percentOfRow'
            ? aggregateBase(`row${rowIndex}`, dataFieldIndex, rowPart.predicates)
            : aggregateBase(`col${colIndex}`, dataFieldIndex, colPart.predicates)
      // Leave blank when the denominator is null or 0 (Excel shows #DIV/0!;
      // the recompute engine opts for blank).
      line.push(base === null || base === 0 ? null : raw / base)
    })
    data.push(line)
  })
  return { data }
}

// ─── Automatic layout growth ────────────────────────────────────────────────

export interface PivotGrowth {
  /// The grown definition; when nothing grew, it is the same object passed in.
  readonly definition: PivotDefinition
  readonly grown: boolean
}

/// Automatic layout growth on refresh: when the source data has categories on
/// axis/report-filter fields that are missing from the cache, refresh is no
/// longer rejected — new members are appended to sharedItems/fieldItems, and
/// the affected axes' layout lines are rebuilt in per-level fieldItems order
/// (old members keep their order, new members are appended at the end of
/// their level, the standard append-on-refresh order).
///
/// Cases that still fail closed (PivotRefreshError thrown as-is):
///  - header mismatch / narrowed source range (the source may have moved);
///  - unsupported definitions (calculated fields, value filters, etc.);
///  - compact/outline layouts (data lines not pinning every level) and
///    layouts containing blank lines.
///
/// TODO(OOXML persistence): growth happens only on the in-memory definition;
/// on save we rely on refreshOnLoad so Excel rebuilds the cache and
/// rowItems/colItems, while the location ref expansion is written back
/// separately by PivotRefreshUpdate. Full write-back of sharedItems +
/// rowItems is left for later.
export function growPivotDefinition(
  definition: PivotDefinition,
  sourceValues: readonly (readonly SourceCell[])[],
): PivotGrowth {
  assertSourceMatchesCache(definition, sourceValues)
  const rows = sourceValues.slice(1)

  // For each axis/filter field: collect members missing from the cache (in
  // first-seen order); for grouped fields the new members are the bucketed
  // group labels (same space as sharedItems).
  const additionsByField = new Map<number, PivotSharedItem[]>()
  for (const field of axisFields(definition)) {
    const grouping = definition.fields[field]?.grouping
    const known = new Set(definition.fields[field]?.sharedItems.map(groupKey) ?? [])
    const additions: PivotSharedItem[] = []
    for (const row of rows) {
      const value = grouping ? groupLabel(grouping, row[field]) : row[field]
      const key = groupKey(value)
      if (known.has(key)) continue
      known.add(key)
      // Blank cells are represented in the cache as <m/> (null).
      additions.push(value === undefined || value === '' ? null : value)
    }
    if (additions.length > 0) additionsByField.set(field, additions)
  }
  // Append to sharedItems and fieldItems (append-only, old indices untouched —
  // item references from layout lines and report filters stay valid).
  let grownBase = definition
  if (additionsByField.size > 0) {
    const fields = definition.fields.map((field, index) => {
      const additions = additionsByField.get(index)
      return additions ? { ...field, sharedItems: [...field.sharedItems, ...additions] } : field
    })
    const fieldItems = definition.fieldItems.map((items, index) => {
      const additions = additionsByField.get(index)
      if (!additions) return items
      const base = definition.fields[index]!.sharedItems.length
      return [...items, ...additions.map((_, offset) => ({ x: base + offset, hidden: false }))]
    })
    grownBase = { ...definition, fields, fieldItems }
  }

  // Reapply value/label filters: filtered-out members are marked as hidden
  // items (this handles both newly grown members and old members that no
  // longer qualify after aggregates changed).
  const reapplied = reapplyPivotFilters(grownBase, rows)
  if (reapplied.changedFields.size > 0) {
    grownBase = { ...grownBase, fieldItems: reapplied.fieldItems }
  }

  if (additionsByField.size === 0 && reapplied.changedFields.size === 0) {
    return { definition, grown: false }
  }

  const affected = (field: number): boolean =>
    additionsByField.has(field) || reapplied.changedFields.has(field)
  const rowAffected = definition.rowFields.some(affected)
  const colAffected = definition.colFields.some(affected)
  return {
    definition: {
      ...grownBase,
      rowLines: rowAffected
        ? rebuildAxisLines(grownBase, definition.rowFields, definition.rowLines, rows)
        : definition.rowLines,
      colLines: colAffected
        ? rebuildAxisLines(grownBase, definition.colFields, definition.colLines, rows)
        : definition.colLines,
    },
    grown: true,
  }
}

// ─── Slicers ─────────────────────────────────────────────────────────────────

/// Slicer-driven member filtering: unselected members of the filtered field
/// are marked as hidden items and the affected axes' layout lines are rebuilt
/// (rows/columns of hidden members vanish from the layout; reselecting
/// restores combinations from the source data). Like label/value filters,
/// hidden items exclude source rows from every aggregate (including
/// subtotals/grand totals) — equivalent to Excel slicers writing hidden items
/// on the pivotField.
///
/// selectedMembers is a set of fieldItems indices; null = all selected (clear
/// the filter). Note: if the field also carries a label/value filter
/// (definition.filters), the next refresh's reapplyPivotFilters recomputes
/// hidden bits from the filter conditions, overriding the slicer's manual
/// selection — slicers and field filters share the same hidden-items
/// state.
export function applyPivotSlicer(
  definition: PivotDefinition,
  sourceValues: readonly (readonly SourceCell[])[],
  field: number,
  selectedMembers: readonly number[] | null,
): PivotDefinition {
  assertSourceMatchesCache(definition, sourceValues)
  if (!axisFields(definition).has(field)) {
    throw new PivotRefreshError(
      'A slicer can only bind to a dimension field in rows/columns/report filters.',
    )
  }
  const items = definition.fieldItems[field] ?? []
  const selected = selectedMembers === null ? null : new Set(selectedMembers)
  if (selected !== null) {
    for (const member of selected) {
      if (items[member] === undefined || items[member]!.x === null) {
        throw new PivotRefreshError(`Slicer-selected member ${member} does not exist.`)
      }
    }
  }
  let changed = false
  const nextItems = items.map((item, member) => {
    if (item.x === null) return item
    const hidden = selected !== null && !selected.has(member)
    if (hidden !== item.hidden) changed = true
    return hidden === item.hidden ? item : { ...item, hidden }
  })
  if (!changed) return definition
  const fieldItems = definition.fieldItems.map((existing, index) =>
    index === field ? nextItems : existing,
  )
  const base: PivotDefinition = { ...definition, fieldItems }
  // Report-filter fields take no row/col layout, so flipping hidden bits is
  // enough; row/col fields need their axis rebuilt.
  const rows = sourceValues.slice(1)
  return {
    ...base,
    rowLines: definition.rowFields.includes(field)
      ? rebuildAxisLines(base, definition.rowFields, definition.rowLines, rows)
      : definition.rowLines,
    colLines: definition.colFields.includes(field)
      ? rebuildAxisLines(base, definition.colFields, definition.colLines, rows)
      : definition.colLines,
  }
}

/// Recompute member visibility for each filtered field per definition.filters:
/// label filters match member labels directly; value filters first aggregate
/// each candidate member (scoped to report filters plus all label filters,
/// without stacking other value filters, avoiding order dependence), then
/// keep members by top/greater-than/between. Returns the updated fieldItems
/// and the set of fields whose hidden bits changed; returns as-is if no filters.
function reapplyPivotFilters(
  definition: PivotDefinition,
  rows: readonly (readonly SourceCell[])[],
): {
  fieldItems: readonly (readonly PivotFieldItem[])[]
  changedFields: Set<number>
} {
  const changedFields = new Set<number>()
  if (definition.filters.length === 0) {
    return { fieldItems: definition.fieldItems, changedFields }
  }

  const itemsOf = (field: number): readonly PivotFieldItem[] => definition.fieldItems[field] ?? []
  const memberLabel = (field: number, item: PivotFieldItem): string => {
    const shared = item.x === null ? undefined : definition.fields[field]?.sharedItems[item.x]
    return shared === null || shared === undefined ? '' : String(shared)
  }
  // Candidate member sets (in fieldItems order), lazily initialized per field
  // to all real items.
  const allowedByField = new Map<number, Set<number>>()
  const ensureAllowed = (field: number): Set<number> => {
    let allowed = allowedByField.get(field)
    if (!allowed) {
      allowed = new Set<number>()
      itemsOf(field).forEach((item, member) => {
        if (item.x !== null) allowed!.add(member)
      })
      allowedByField.set(field, allowed)
    }
    return allowed
  }

  // Phase one: label filters.
  for (const filter of definition.filters) {
    if (filter.kind !== 'label') continue
    const allowed = ensureAllowed(filter.field)
    itemsOf(filter.field).forEach((item, member) => {
      if (item.x === null || !allowed.has(member)) return
      if (!matchesLabelFilter(filter, memberLabel(filter.field, item))) allowed.delete(member)
    })
  }

  // Phase two: value filters. Row scope = report filters plus the candidate
  // sets left after all label filters.
  const pagePredicates: AxisPredicate[] = []
  for (const page of definition.pageFields) {
    if (page.item === null) continue
    const item = definition.fieldItems[page.field]?.[page.item]
    const shared =
      item?.x === null || item === undefined
        ? undefined
        : definition.fields[page.field]?.sharedItems[item.x]
    if (shared === undefined)
      throw new PivotRefreshError('The selected report filter item cannot be resolved.')
    pagePredicates.push({ field: page.field, key: groupKey(shared) })
  }
  // Member key → member index (within candidate sets), used to assign source
  // rows to members.
  const memberKeys = new Map<number, Map<string, number>>()
  const ensureMemberKeys = (field: number): Map<string, number> => {
    let keys = memberKeys.get(field)
    if (!keys) {
      keys = new Map<string, number>()
      itemsOf(field).forEach((item, member) => {
        if (item.x === null) return
        const shared = definition.fields[field]?.sharedItems[item.x]
        if (shared !== undefined) keys!.set(groupKey(shared), member)
      })
      memberKeys.set(field, keys)
    }
    return keys
  }
  const rowPassesBase = (row: readonly SourceCell[]): boolean => {
    for (const page of pagePredicates) {
      if (sourceKey(definition, page.field, row[page.field]) !== page.key) return false
    }
    for (const [field, allowed] of allowedByField) {
      const member = ensureMemberKeys(field).get(sourceKey(definition, field, row[field]))
      if (member === undefined || !allowed.has(member)) return false
    }
    return true
  }
  for (const filter of definition.filters) {
    if (filter.kind !== 'value') continue
    const dataField = definition.dataFields[filter.dataField]
    if (!dataField)
      throw new PivotRefreshError(
        `The data field ${filter.dataField} referenced by the value filter is missing.`,
      )
    const allowed = ensureAllowed(filter.field)
    const keys = ensureMemberKeys(filter.field)
    const baseRows = rows.filter(rowPassesBase)
    const totals = [...allowed]
      .sort((a, b) => a - b)
      .map((member) => ({
        key: member,
        total: aggregateDataFieldOverRows(
          definition,
          dataField,
          baseRows.filter(
            (row) => keys.get(sourceKey(definition, filter.field, row[filter.field])) === member,
          ),
        ),
      }))
    const kept = allowedByValueFilter(filter, totals)
    for (const member of [...allowed]) {
      if (!kept.has(member)) allowed.delete(member)
    }
  }

  // Apply: a filtered field's hidden bit = not in its candidate set (items
  // outside the data-field axis stay untouched).
  const fieldItems = definition.fieldItems.map((items, field) => {
    const allowed = allowedByField.get(field)
    if (!allowed) return items
    let changed = false
    const next = items.map((item, member) => {
      if (item.x === null) return item
      const hidden = !allowed.has(member)
      if (hidden !== item.hidden) changed = true
      return hidden === item.hidden ? item : { ...item, hidden }
    })
    if (!changed) return items
    changedFields.add(field)
    return next
  })
  return { fieldItems, changedFields }
}

/// Aggregate one data field over a set of source rows: plain fields use their
/// subtotal; calculated fields evaluate the formula over each referenced
/// field's SUM (same semantics as recomputePivotData).
function aggregateDataFieldOverRows(
  definition: PivotDefinition,
  dataField: PivotDataField,
  rows: readonly (readonly SourceCell[])[],
): number | null {
  if (dataField.formula !== undefined) {
    const indexByName = new Map<string, number>()
    definition.fields.forEach((field, index) => {
      if (field.formula === undefined) indexByName.set(field.name, index)
    })
    try {
      const ast = parsePivotFormula(dataField.formula, [...indexByName.keys()])
      return evaluatePivotFormula(ast, (name) =>
        aggregate(
          'sum',
          rows.map((row) => row[indexByName.get(name)!]),
        ),
      )
    } catch (error) {
      throw new PivotRefreshError(
        `The formula of calculated field "${dataField.name}" cannot be parsed: ` +
          `${error instanceof Error ? error.message : String(error)}`,
      )
    }
  }
  return aggregate(
    dataField.subtotal,
    rows.map((row) => row[dataField.field]),
  )
}

/// Rebuild one axis's layout lines: expand actually-present member
/// combinations level by level in fieldItems order (old data-line combos ∪
/// source-data combos), keeping the original layout's subtotal depths and
/// grand-total lines; the data-field axis level (-2) expands one line per
/// data field.
function rebuildAxisLines(
  definition: PivotDefinition,
  axis: readonly number[],
  oldLines: readonly PivotLayoutLine[],
  rows: readonly (readonly SourceCell[])[],
): PivotLayoutLine[] {
  if (oldLines.some((line) => line.t === 'blank')) {
    throw new PivotRefreshError(
      'Pivot layouts containing blank lines do not support automatic growth — refresh in Excel.',
    )
  }
  const levels = axis.length
  // Compact/outline layouts have data lines that don't pin every level
  // (parent heading lines); rebuilding would lose that structure — fail-closed.
  if (oldLines.some((line) => line.t === 'data' && line.depth !== levels)) {
    throw new PivotRefreshError(
      'Compact/outline pivot layouts do not support automatic growth yet — refresh in Excel.',
    )
  }

  // Display order and member keys per level; on the -2 (data-field axis)
  // level the members are data-field indices.
  const orderPerLevel = axis.map((field) => {
    if (field === -2) {
      return definition.dataFields.map((_, index) => ({ member: index, key: `d${index}` }))
    }
    const entries: { member: number; key: string }[] = []
    definition.fieldItems[field]?.forEach((item, member) => {
      if (item.hidden || item.x === null) return
      const shared = definition.fields[field]?.sharedItems[item.x]
      if (shared !== undefined) entries.push({ member, key: groupKey(shared) })
    })
    return entries
  })

  // Combination-prefix set over real field levels (excluding -2); keys join
  // with NUL (\u0000) to avoid ambiguity.
  const realLevels = axis
    .map((field, level) => ({ field, level }))
    .filter((entry) => entry.field !== -2)
  const prefixes = new Set<string>()
  const joinKeys = (keys: readonly string[]): string => keys.join('\u0000')
  const addCombo = (keys: readonly string[]): void => {
    for (let length = 1; length <= keys.length; length += 1) {
      prefixes.add(joinKeys(keys.slice(0, length)))
    }
  }
  const memberKeys = new Map<string, string>()
  orderPerLevel.forEach((entries, level) => {
    const field = axis[level]!
    if (field === -2) return
    for (const entry of entries) memberKeys.set(`${field}:${entry.member}`, entry.key)
  })
  for (const line of oldLines) {
    if (line.t !== 'data') continue
    const keys: string[] = []
    let resolvable = true
    for (const { field, level } of realLevels) {
      const member = line.members[level]
      const key =
        member === null || member === undefined ? undefined : memberKeys.get(`${field}:${member}`)
      if (key === undefined) {
        resolvable = false
        break
      }
      keys.push(key)
    }
    // Old lines referencing hidden/stale items: their combos are dropped
    // (same result as a refresh).
    if (resolvable) addCombo(keys)
  }
  for (const row of rows) {
    addCombo(realLevels.map(({ field }) => sourceKey(definition, field, row[field])))
  }

  const subtotalDepths = new Set(
    oldLines.filter((line) => line.t === 'default').map((line) => line.depth),
  )
  const rebuilt: PivotLayoutLine[] = []
  const emit = (
    level: number,
    members: readonly (number | null)[],
    realKeys: readonly string[],
    dataField: number,
  ): void => {
    const field = axis[level]!
    for (const entry of orderPerLevel[level]!) {
      const nextKeys = field === -2 ? realKeys : [...realKeys, entry.key]
      if (field !== -2 && !prefixes.has(joinKeys(nextKeys))) continue
      const nextMembers = members.map((member, index) => (index === level ? entry.member : member))
      const nextDataField = field === -2 ? entry.member : dataField
      if (level === levels - 1) {
        rebuilt.push({ t: 'data', members: nextMembers, depth: levels, dataField: nextDataField })
      } else {
        emit(level + 1, nextMembers, nextKeys, nextDataField)
        if (subtotalDepths.has(level + 1)) {
          rebuilt.push({
            t: 'default',
            members: nextMembers.map((member, index) => (index <= level ? member : null)),
            depth: level + 1,
            dataField: nextDataField,
          })
        }
      }
    }
  }
  if (levels > 0) emit(0, new Array(levels).fill(null), [], 0)
  // Grand-total lines are kept as-is (on a data-field axis there may be one
  // per data field).
  for (const line of oldLines) {
    if (line.t === 'grand') rebuilt.push(line)
  }
  return rebuilt
}
