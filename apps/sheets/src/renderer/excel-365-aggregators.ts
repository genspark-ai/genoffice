/**
 * Excel 365's GROUPBY and PIVOTBY. Univer has neither executor, and its
 * lexer cannot hand a bare function name (GROUPBY's third argument) into a
 * calculation — the name parses as an undefined name and reaches the
 * executor as a #NAME? error with the original spelling gone. An Excel
 * workbook using them recalculates to #NAME?.
 *
 * Two pieces:
 *
 * - a lexer wrap that rewrites the function-argument position of
 *   GROUPBY/PIVOTBY into a string literal ("SUM") before the tree builder
 *   sees it. The rewritten node keeps its original indices, so every
 *   consumer that reassembles the formula from the sequence (the editor's
 *   normalization, autofill offsets, F4, the saved cell text) still reads
 *   the bare name Excel wrote — only the calculation sees the string.
 * - the executors: group the row fields, aggregate each values column
 *   through the named executor (fetched live from the function service),
 *   and lay the optional headers, totals and sort order over the result.
 *
 * Knives left for later: total_depth 2 (a subtotal row per group) and the
 * negative depths (totals at top), a HSTACK(SUM,AVERAGE) function pair,
 * multi-column col_fields in PIVOTBY (hierarchical column headers), and the
 * aggregate-value sort encodings. field_headers=2 generates positional
 * labels ("Field 1", "Value 1") — Excel derives them from the source
 * table's field names, which a bare array does not carry.
 */
import {
  ArrayValueObject,
  BaseFunction,
  type BaseValueObject,
  ErrorType,
  ErrorValueObject,
  FunctionType,
  IFunctionService,
  LexerTreeBuilder,
  StringValueObject,
  type IFunctionInfo,
} from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

/** the argument position that carries the aggregating function, 0-based */
const GROUP_FN_ARG: Record<string, number> = { GROUPBY: 2, PIVOTBY: 3 }

/** the label a total row carries (Excel's English literal; a locale-following
    label is a later knife) */
const TOTAL_LABEL = 'Total'

/** a bare identifier that could name a function: caps, digits and dots */
const BARE_FN_NAME = /^[A-Z][A-Z0-9.]*$/

interface LexerNodeLike {
  getToken(): string
  setToken(token: string): void
  getChildren(): (LexerNodeLike | string)[]
  setChildren(children: (LexerNodeLike | string)[]): void
  getParent(): LexerNodeLike | null | undefined
}

interface LexerInternals {
  _nodeMaker: (formula: string, sequenceArray?: unknown, matchCurrentNodeIndex?: number) => unknown
  _currentLexerNode: LexerNodeLike
}

/**
 * Walk the parsed tree and replace the bare function name in GROUPBY/PIVOTBY's
 * function-argument position with a string literal the executor reads. The
 * function node's children are one P-node per argument, and a bare name
 * argument is a single string child — a call keeps its own P-node subtree.
 * Rewriting the tree leaves the formula text untouched: the editor, autofill
 * and the saved cell text still carry the bare name Excel wrote.
 */
function rewriteTree(node: LexerNodeLike, functionService: IFunctionService): void {
  const children = node.getChildren()
  const fnArgIndex = GROUP_FN_ARG[node.getToken()]
  if (fnArgIndex !== undefined) {
    const param = children[fnArgIndex]
    if (param && typeof param !== 'string') {
      const inner = param.getChildren()
      if (inner.length === 1 && typeof inner[0] === 'string') {
        const name = inner[0]!
        if (BARE_FN_NAME.test(name) && functionService.getExecutor(name) != null) {
          param.setChildren([`"${name}"`])
        }
      }
    }
  }
  for (const child of children) {
    if (typeof child !== 'string') rewriteTree(child, functionService)
  }
}

/**
 * The wrap over the injector's LexerTreeBuilder singleton — the same instance
 * the editor, autofill and the engine all resolve (see formula-lexer-fix for
 * the precedent). Every full parse of a GROUPBY/PIVOTBY formula gets the
 * rewrite; the cursor-position parse the editor runs for its own bookkeeping
 * is left alone.
 */
export function wrapGroupByFunctionRefs(
  lexer: LexerTreeBuilder,
  functionService: IFunctionService,
): { dispose(): void } {
  const internals = lexer as unknown as LexerInternals
  const original = internals._nodeMaker
  internals._nodeMaker = function (
    this: LexerInternals,
    formula: string,
    sequenceArray?: unknown,
    matchIndex?: number,
  ) {
    const result = original.call(this, formula, sequenceArray, matchIndex)
    if (matchIndex !== undefined || !/GROUPBY|PIVOTBY/.test(formula)) return result
    let top = this._currentLexerNode
    while (top?.getParent()) top = top.getParent() as LexerNodeLike
    if (top) rewriteTree(top, functionService)
    return result
  }
  return {
    dispose() {
      internals._nodeMaker = original
    },
  }
}

/** an optional numeric argument: missing → default, bad → the error to return */
function numberArg(
  variant: BaseValueObject | undefined,
  fallback: number,
): { value: number } | { error: BaseValueObject } {
  if (variant === undefined || variant.isNull?.()) return { value: fallback }
  if (variant.isError?.()) return { error: variant }
  const raw = variant.getValue()
  if (typeof raw !== 'number') return { error: errorOf(ErrorType.VALUE) }
  return { value: raw }
}

function errorOf(code: ErrorType): ErrorValueObject {
  return ErrorValueObject.create(code)
}

/** the engine's isError is a runtime check, not a type predicate */
function isError(value: BaseValueObject): value is ErrorValueObject {
  return value.isError?.() === true
}

function asArray(value: BaseValueObject): ArrayValueObject | ErrorValueObject {
  if (value.isError?.()) return value as ErrorValueObject
  return value as ArrayValueObject
}

/** a stable text key for one row of the grouping fields */
function keyTextOf(rowFields: ArrayValueObject, row: number, keyColumns: number): string {
  let text = ''
  for (let column = 0; column < keyColumns; column++) {
    const cell = rowFields.get(row, column)
    text += `\u0001${cell?.getValue() ?? ''}`
  }
  return text
}

/** the boolean mask a filter_array carries, aligned to the field rows */
function readFilter(filter: ArrayValueObject, rowCount: number): boolean[] {
  const out: boolean[] = []
  for (let row = 0; row < rowCount; row++) {
    const value = filter.get(row, 0)?.getValue()
    if (typeof value === 'boolean') out.push(value)
    else if (typeof value === 'number') out.push(value !== 0)
    else out.push(false)
  }
  return out
}

/** an output array from a cell-object matrix (errors stay errors, numbers
    keep their type) */
function arrayOf(cells: (BaseValueObject | null)[][]): ArrayValueObject {
  return ArrayValueObject.create({
    calculateValueList: cells,
    rowCount: cells.length,
    columnCount: cells[0]?.length ?? 0,
    unitId: '',
    sheetId: '',
    row: -1,
    column: -1,
  })
}

interface Grouped {
  /** the distinct key rows, in output order, as cell objects */
  keys: BaseValueObject[][]
  /** the data rows behind each key */
  rows: number[][]
}

function groupRows(
  rowFields: ArrayValueObject,
  dataRows: readonly number[],
  keyColumns: number,
): Grouped {
  const keys: BaseValueObject[][] = []
  const rows: number[][] = []
  const at = new Map<string, number>()
  for (const row of dataRows) {
    const text = keyTextOf(rowFields, row, keyColumns)
    let index = at.get(text)
    if (index === undefined) {
      index = keys.length
      at.set(text, index)
      keys.push(
        Array.from(
          { length: keyColumns },
          (_, column) => rowFields.get(row, column) ?? StringValueObject.create(''),
        ),
      )
      rows.push([])
    }
    rows[index]!.push(row)
  }
  return { keys, rows }
}

/** order the groups: sort_order names a row-field column (1-based, negative
    for descending); 0 keeps first-seen order */
function orderGroups(grouped: Grouped, sortOrder: number, keyColumns: number): Grouped {
  if (sortOrder === 0 || grouped.keys.length <= 1) return grouped
  const column = Math.abs(sortOrder) - 1
  const direction = sortOrder > 0 ? 1 : -1
  if (column >= keyColumns) return grouped
  const order = grouped.keys
    .map((key, index) => ({ raw: key[column]?.getValue(), index }))
    .sort((a, b) => {
      if (typeof a.raw === 'number' && typeof b.raw === 'number') return (a.raw - b.raw) * direction
      return String(a.raw ?? '').localeCompare(String(b.raw ?? '')) * direction
    })
    .map((entry) => entry.index)
  return { keys: order.map((at) => grouped.keys[at]!), rows: order.map((at) => grouped.rows[at]!) }
}

/**
 * The field_headers argument, Excel's own encoding (support.microsoft.com
 * GROUPBY): 0 no headers, 1 headers exist but stay hidden, 2 no headers but
 * generate them, 3 headers exist and show. Missing — and anything out of
 * range — is Excel's "automatic", which decides from the values argument
 * alone (headers when the first value is text and the second a number); the
 * row fields are never consulted.
 */
function decodeFieldHeaders(
  headers: number,
  values: ArrayValueObject,
): { have: boolean; show: boolean; generate: boolean } {
  switch (headers) {
    case 0:
      return { have: false, show: false, generate: false }
    case 1:
      return { have: true, show: false, generate: false }
    case 2:
      return { have: false, show: true, generate: true }
    case 3:
      return { have: true, show: true, generate: false }
    default: {
      const first = values.get(0, 0)?.getValue()
      const second = values.get(1, 0)?.getValue()
      const have = typeof first === 'string' && typeof second === 'number'
      return { have, show: have, generate: false }
    }
  }
}

/** the header row field_headers=2 generates; Excel derives these from the
    source table's field names, which a bare array does not carry */
const generatedHeaders = (keyColumns: number, valueColumns: number): string[] => [
  ...Array.from({ length: keyColumns }, (_, i) => `Field ${i + 1}`),
  ...Array.from({ length: valueColumns }, (_, i) => `Value ${i + 1}`),
]

/** aggregate `values` column `valueColumn` over the given rows */
function aggregateRows(
  values: ArrayValueObject,
  rows: readonly number[],
  valueColumn: number,
  fnName: string,
  functionService: IFunctionService,
): BaseValueObject {
  const executor = functionService.getExecutor(fnName)
  if (!executor) return errorOf(ErrorType.NAME)
  const slice = rows.map((row) => [values.get(row, valueColumn) ?? null])
  return executor.calculate(
    ArrayValueObject.create({
      calculateValueList: slice,
      rowCount: Math.max(slice.length, 1),
      columnCount: 1,
      unitId: '',
      sheetId: '',
      row: -1,
      column: -1,
    }),
  ) as BaseValueObject
}

function textValue(value: string): StringValueObject {
  return StringValueObject.create(value)
}

/**
 * GROUPBY(row_fields, values, function, [field_headers], [total_depth],
 * [sort_order], [filter_array])
 */
class GroupByFunction extends BaseFunction {
  private readonly _functionService: IFunctionService

  constructor(functionService: IFunctionService) {
    super('GROUPBY')
    this._functionService = functionService
  }

  override calculate(...variants: BaseValueObject[]): BaseValueObject {
    if (variants.length < 3) return errorOf(ErrorType.VALUE)
    const rowFields = asArray(variants[0]!)
    const values = asArray(variants[1]!)
    if (isError(rowFields)) return rowFields
    if (isError(values)) return values
    const fnName = variants[2]?.getValue()
    if (typeof fnName !== 'string') return errorOf(ErrorType.VALUE)
    // missing field_headers is Excel's automatic, decided from the values
    // argument alone; missing total_depth is automatic = grand totals;
    // missing sort_order is Excel's default ascending by the first row field
    const headers = numberArg(variants[3], -1)
    if ('error' in headers) return headers.error
    const totalDepth = numberArg(variants[4], 1)
    if ('error' in totalDepth) return totalDepth.error
    const sortOrder = numberArg(variants[5], 1)
    if ('error' in sortOrder) return sortOrder.error
    const filterVariant = variants[6]
    const rowCount = Math.min(rowFields.getRowCount(), values.getRowCount())
    const headerMode = decodeFieldHeaders(headers.value, values)
    if (headerMode.have && rowCount < 2) return errorOf(ErrorType.VALUE)
    const dataStart = headerMode.have ? 1 : 0
    let dataRows: number[] = []
    for (let row = dataStart; row < rowCount; row++) dataRows.push(row)
    if (filterVariant && !filterVariant.isNull?.()) {
      const filter = asArray(filterVariant)
      if (isError(filter)) return filter
      const mask = readFilter(filter, rowCount)
      dataRows = dataRows.filter((row) => mask[row])
    }
    const grouped = groupRows(rowFields, dataRows, rowFields.getColumnCount())
    const ordered = orderGroups(grouped, sortOrder.value, rowFields.getColumnCount())
    const keyColumns = rowFields.getColumnCount()
    const valueColumns = values.getColumnCount()
    const out: (BaseValueObject | null)[][] = []
    if (headerMode.show) {
      const head: (BaseValueObject | null)[] = []
      const generated = headerMode.generate ? generatedHeaders(keyColumns, valueColumns) : null
      for (let column = 0; column < keyColumns; column++) {
        head.push(generated ? textValue(generated[column]!) : (rowFields.get(0, column) ?? null))
      }
      for (let column = 0; column < valueColumns; column++) {
        head.push(
          generated ? textValue(generated[keyColumns + column]!) : (values.get(0, column) ?? null),
        )
      }
      out.push(head)
    }
    for (let group = 0; group < ordered.keys.length; group++) {
      const row: (BaseValueObject | null)[] = [...ordered.keys[group]!]
      for (let column = 0; column < valueColumns; column++) {
        row.push(aggregateRows(values, ordered.rows[group]!, column, fnName, this._functionService))
      }
      out.push(row)
    }
    if (totalDepth.value > 0 && ordered.keys.length > 0) {
      const total: (BaseValueObject | null)[] = []
      for (let column = 0; column < keyColumns; column++) {
        total.push(column === 0 ? textValue(TOTAL_LABEL) : null)
      }
      for (let column = 0; column < valueColumns; column++) {
        total.push(aggregateRows(values, dataRows, column, fnName, this._functionService))
      }
      out.push(total)
    }
    if (out.length === 0) return errorOf(ErrorType.VALUE)
    return arrayOf(out)
  }
}

/**
 * PIVOTBY(row_fields, col_fields, values, function, [field_headers],
 * [row_total_depth], [row_sort_order], [col_total_depth],
 * [col_sort_order], [filter_array]) — one column field, grand totals on
 * both axes.
 */
class PivotByFunction extends BaseFunction {
  private readonly _functionService: IFunctionService

  constructor(functionService: IFunctionService) {
    super('PIVOTBY')
    this._functionService = functionService
  }

  override calculate(...variants: BaseValueObject[]): BaseValueObject {
    if (variants.length < 4) return errorOf(ErrorType.VALUE)
    const rowFields = asArray(variants[0]!)
    const colFields = asArray(variants[1]!)
    const values = asArray(variants[2]!)
    if (isError(rowFields)) return rowFields
    if (isError(colFields)) return colFields
    if (isError(values)) return values
    if (colFields.getColumnCount() > 1) return errorOf(ErrorType.VALUE)
    const fnName = variants[3]?.getValue()
    if (typeof fnName !== 'string') return errorOf(ErrorType.VALUE)
    // field_headers: Excel's encoding, automatic from the values argument
    // when missing — same decode as GROUPBY
    const headers = numberArg(variants[4], -1)
    if ('error' in headers) return headers.error
    const rowTotal = numberArg(variants[5], 1)
    if ('error' in rowTotal) return rowTotal.error
    const rowSort = numberArg(variants[6], 0)
    if ('error' in rowSort) return rowSort.error
    const colTotal = numberArg(variants[7], 1)
    if ('error' in colTotal) return colTotal.error
    const colSort = numberArg(variants[8], 0)
    if ('error' in colSort) return colSort.error
    const filterVariant = variants[9]
    const rowCount = Math.min(
      rowFields.getRowCount(),
      colFields.getRowCount(),
      values.getRowCount(),
    )
    const headerMode = decodeFieldHeaders(headers.value, values)
    if (headerMode.have && rowCount < 2) return errorOf(ErrorType.VALUE)
    const dataStart = headerMode.have ? 1 : 0
    let dataRows: number[] = []
    for (let row = dataStart; row < rowCount; row++) dataRows.push(row)
    if (filterVariant && !filterVariant.isNull?.()) {
      const filter = asArray(filterVariant)
      if (isError(filter)) return filter
      const mask = readFilter(filter, rowCount)
      dataRows = dataRows.filter((row) => mask[row])
    }
    const keyColumns = rowFields.getColumnCount()
    // distinct row keys and column keys, in first-seen order
    const grouped = groupRows(rowFields, dataRows, keyColumns)
    const colKeys: BaseValueObject[] = []
    const colRows: number[][] = []
    const colAt = new Map<string, number>()
    for (const row of dataRows) {
      const cell = colFields.get(row, 0)
      const text = `\u0001${cell?.getValue() ?? ''}`
      let index = colAt.get(text)
      if (index === undefined) {
        index = colKeys.length
        colAt.set(text, index)
        colKeys.push(cell ?? textValue(''))
        colRows.push([])
      }
      colRows[index]!.push(row)
    }
    const orderedRows = orderGroups(grouped, rowSort.value, keyColumns)
    // the column keys order: a negative col_sort_order reverses first-seen
    let columnOrder = colKeys.map((unused, index) => index)
    if (colSort.value !== 0 && colKeys.length > 1) {
      columnOrder = colSort.value > 0 ? columnOrder : [...columnOrder].reverse()
    }
    const out: (BaseValueObject | null)[][] = []
    if (headerMode.show) {
      const head: (BaseValueObject | null)[] = []
      const generated = headerMode.generate ? generatedHeaders(keyColumns, 0) : null
      for (let column = 0; column < keyColumns; column++) {
        head.push(generated ? textValue(generated[column]!) : (rowFields.get(0, column) ?? null))
      }
      for (const index of columnOrder) head.push(colKeys[index]!)
      if (rowTotal.value > 0) head.push(textValue(TOTAL_LABEL))
      out.push(head)
    }
    for (let group = 0; group < orderedRows.keys.length; group++) {
      const row: (BaseValueObject | null)[] = [...orderedRows.keys[group]!]
      const rows = new Set(orderedRows.rows[group]!)
      for (const index of columnOrder) {
        const slice = colRows[index]!.filter((at) => rows.has(at))
        row.push(
          slice.length > 0 ? aggregateRows(values, slice, 0, fnName, this._functionService) : null,
        )
      }
      if (rowTotal.value > 0) {
        row.push(aggregateRows(values, orderedRows.rows[group]!, 0, fnName, this._functionService))
      }
      out.push(row)
    }
    if (colTotal.value > 0 && dataRows.length > 0) {
      const total: (BaseValueObject | null)[] = []
      for (let column = 0; column < keyColumns; column++) {
        total.push(column === 0 ? textValue(TOTAL_LABEL) : null)
      }
      for (const index of columnOrder) {
        total.push(aggregateRows(values, colRows[index]!, 0, fnName, this._functionService))
      }
      if (rowTotal.value > 0) {
        total.push(aggregateRows(values, dataRows, 0, fnName, this._functionService))
      }
      out.push(total)
    }
    if (out.length === 0) return errorOf(ErrorType.VALUE)
    return arrayOf(out)
  }
}

const EXCEL_365_AGGREGATOR_DESCRIPTIONS: IFunctionInfo[] = [
  {
    functionName: 'GROUPBY',
    functionType: FunctionType.Database,
    abstract: 'Groups a range by its row fields and aggregates each group',
    description:
      'Groups the rows of a range by the row fields, aggregates each group with the named function, and lays out the groups with optional headers, totals and sort order.',
    functionParameter: [
      {
        name: 'row_fields',
        detail: 'The columns to group by',
        example: 'A1:A10',
        require: 1,
        repeat: 0,
      },
      {
        name: 'values',
        detail: 'The columns to aggregate',
        example: 'B1:B10',
        require: 1,
        repeat: 0,
      },
      {
        name: 'function',
        detail: 'The aggregating function, e.g. SUM',
        example: 'SUM',
        require: 1,
        repeat: 0,
      },
      {
        name: 'field_headers',
        detail: '0 no headers, 1 headers, 2 auto',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'total_depth',
        detail: '0 no totals, 1 grand total',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'sort_order',
        example: '',
        detail: 'Row-field column to sort by; negative for descending',
        require: 0,
        repeat: 0,
      },
      {
        name: 'filter_array',
        detail: 'A boolean mask over the data rows',
        example: '',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'PIVOTBY',
    functionType: FunctionType.Database,
    abstract: 'Pivots a range by row and column fields and aggregates each cell',
    description:
      'Groups the rows of a range by the row fields, spreads the column field across the output columns, aggregates each cell with the named function, and adds optional grand totals.',
    functionParameter: [
      {
        name: 'row_fields',
        detail: 'The columns to group rows by',
        example: '',
        require: 1,
        repeat: 0,
      },
      {
        name: 'col_fields',
        detail: 'The column to spread across',
        example: 'B1:B10',
        require: 1,
        repeat: 0,
      },
      { name: 'values', detail: 'The column to aggregate', example: '', require: 1, repeat: 0 },
      {
        name: 'function',
        detail: 'The aggregating function, e.g. SUM',
        example: 'SUM',
        require: 1,
        repeat: 0,
      },
      {
        name: 'field_headers',
        detail: '0 no headers, 1 headers, 2 auto',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'row_total_depth',
        detail: '0 no row totals, 1 grand total',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'row_sort_order',
        example: '',
        detail: 'Row-field column to sort by; negative for descending',
        require: 0,
        repeat: 0,
      },
      {
        name: 'col_total_depth',
        detail: '0 no column totals, 1 grand total',
        example: '1',
        require: 0,
        repeat: 0,
      },
      {
        name: 'col_sort_order',
        detail: 'Negative for descending columns',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'filter_array',
        detail: 'A boolean mask over the data rows',
        example: '',
        require: 0,
        repeat: 0,
      },
    ],
  },
]

export function installExcel365Aggregators(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const functionService = injector.get(IFunctionService)
  functionService.registerExecutors(
    new GroupByFunction(functionService),
    new PivotByFunction(functionService),
  )
  const descriptions = functionService.registerDescriptions(...EXCEL_365_AGGREGATOR_DESCRIPTIONS)
  const unwrapLexer = wrapGroupByFunctionRefs(injector.get(LexerTreeBuilder), functionService)
  return {
    dispose() {
      unwrapLexer.dispose()
      descriptions.dispose()
      functionService.unregisterExecutors('GROUPBY', 'PIVOTBY')
    },
  }
}
