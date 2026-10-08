/**
 * Pure helpers behind Excel's Function Arguments dialog: locating the call
 * under the caret, splitting/joining argument lists, laying out one row per
 * parameter (repeating parameters grow as they fill), rendering preview
 * values and keeping the Most Recently Used list.
 */

export interface ParamInfo {
  readonly name: string
  readonly detail: string
  readonly require: boolean
  readonly repeat: boolean
}

export interface FunctionCallSpan {
  readonly name: string
  /// Index of the first character of the name.
  readonly start: number
  /// Index after the closing paren, or the formula's length when unclosed.
  readonly end: number
  readonly args: readonly string[]
}

export interface ArgumentRow {
  readonly name: string
  readonly detail: string
  readonly require: boolean
  readonly index: number
}

/// Advances past a string literal ("" escapes), a quoted sheet name, or an
/// array literal; returns the index of the character following it.
function skipLiteral(text: string, i: number): number {
  const open = text[i]
  if (open === '"' || open === "'") {
    for (let j = i + 1; j < text.length; j++) {
      if (text[j] !== open) continue
      if (text[j + 1] === open) {
        j++
        continue
      }
      return j + 1
    }
    return text.length
  }
  if (open === '{') {
    const close = text.indexOf('}', i + 1)
    return close === -1 ? text.length : close + 1
  }
  return i + 1
}

/// Splits the text between a call's parens at top-level commas; literals and
/// nested calls stay intact. Whitespace around each argument is trimmed.
export function splitTopLevelArgs(text: string): string[] {
  if (text.trim() === '') return []
  const args: string[] = []
  let depth = 0
  let from = 0
  let i = 0
  while (i < text.length) {
    const ch = text[i]!
    if (ch === '"' || ch === "'" || ch === '{') {
      i = skipLiteral(text, i)
      continue
    }
    if (ch === '(') depth++
    else if (ch === ')') depth--
    else if (ch === ',' && depth === 0) {
      args.push(text.slice(from, i).trim())
      from = i + 1
    }
    i++
  }
  args.push(text.slice(from).trim())
  return args
}

const NAME_TAIL = /[A-Za-z0-9_.]/

/// The innermost function call whose parens hold the caret (a caret right
/// after the closing paren still counts, so `=SUM(A1:A3)|` opens SUM).
export function findFunctionCallAtCaret(formula: string, caret: number): FunctionCallSpan | null {
  const stack: { nameStart: number; open: number }[] = []
  const calls: FunctionCallSpan[] = []
  let i = 0
  while (i < formula.length) {
    const ch = formula[i]!
    if (ch === '"' || ch === "'" || ch === '{') {
      i = skipLiteral(formula, i)
      continue
    }
    if (ch === '(') {
      let nameStart = i
      while (nameStart > 0 && NAME_TAIL.test(formula[nameStart - 1]!)) nameStart--
      if (nameStart < i && /[A-Za-z_]/.test(formula[nameStart]!)) stack.push({ nameStart, open: i })
      else stack.push({ nameStart: -1, open: i })
    } else if (ch === ')') {
      const frame = stack.pop()
      if (frame && frame.nameStart >= 0) {
        calls.push({
          name: formula.slice(frame.nameStart, frame.open).toUpperCase(),
          start: frame.nameStart,
          end: i + 1,
          args: splitTopLevelArgs(formula.slice(frame.open + 1, i)),
        })
      }
    }
    i++
  }
  for (const frame of stack) {
    if (frame.nameStart < 0) continue
    calls.push({
      name: formula.slice(frame.nameStart, frame.open).toUpperCase(),
      start: frame.nameStart,
      end: formula.length,
      args: splitTopLevelArgs(formula.slice(frame.open + 1)),
    })
  }
  const inside = calls.filter((call) => caret >= call.start && caret <= call.end)
  if (inside.length === 0) return null
  // Innermost = the shortest span that holds the caret
  return inside.reduce((best, call) =>
    call.end - call.start < best.end - best.start ? call : best,
  )
}

/// `NAME(a, b)` with trailing empty arguments dropped; inner empties stay so
/// positional arguments keep their slots.
export function buildFunctionCall(name: string, args: readonly string[]): string {
  const trimmed = args.map((arg) => arg.trim())
  while (trimmed.length > 0 && trimmed[trimmed.length - 1] === '') trimmed.pop()
  return `${name}(${trimmed.join(', ')})`
}

export function spliceFunctionCall(
  formula: string,
  span: Pick<FunctionCallSpan, 'start' | 'end'>,
  call: string,
): string {
  return formula.slice(0, span.start) + call + formula.slice(span.end)
}

/// The call replaces the one under the caret, lands at the caret of a
/// formula without one, or becomes the whole formula of a plain cell.
export function finishedFormula(
  target: { readonly formula: string; readonly caret: number },
  span: Pick<FunctionCallSpan, 'start' | 'end'> | null,
  name: string,
  args: readonly string[],
): string {
  const call = buildFunctionCall(name, args)
  if (target.formula === '') return `=${call}`
  if (span) return spliceFunctionCall(target.formula, span, call)
  return target.formula.slice(0, target.caret) + call + target.formula.slice(target.caret)
}

/// Parameters from an Excel syntax line such as `SUM(number1, [number2], …)`;
/// used for the app's own executors, which the engine does not describe.
export function parseSyntaxParams(syntax: string): ParamInfo[] {
  const open = syntax.indexOf('(')
  const close = syntax.lastIndexOf(')')
  if (open === -1 || close <= open) return []
  const parts = syntax
    .slice(open + 1, close)
    .split(',')
    .map((part) => part.trim())
    .filter((part) => part !== '')
  const repeat = parts[parts.length - 1] === '…' || parts[parts.length - 1] === '...'
  if (repeat) parts.pop()
  return parts.map((part, index) => {
    const optional = part.startsWith('[') && part.endsWith(']')
    return {
      name: optional ? part.slice(1, -1) : part,
      detail: '',
      require: !optional,
      repeat: repeat && index === parts.length - 1,
    }
  })
}

const MAX_ARGUMENT_ROWS = 255

/// Excel's continuing names: `number2` → `number3`, otherwise a numbered copy.
function repeatedName(name: string, ordinal: number): string {
  const match = /^(.*?)(\d+)$/.exec(name)
  if (!match) return `${name}${ordinal}`
  return `${match[1]}${Number(match[2]) + ordinal - 1}`
}

/**
 * One row per parameter. The trailing repeating parameters form a group
 * (SUMIFS repeats criteria_range2 + criteria2 together) that shows as many
 * repetitions as the arguments fill plus one empty repetition to grow into
 * (Excel adds number3 once number2 is filled); only the first repetition
 * keeps its required flags.
 */
export function parameterRows(params: readonly ParamInfo[], argCount: number): ArgumentRow[] {
  const firstRepeat = params.findIndex((param) => param.repeat)
  const fixed = firstRepeat === -1 ? params : params.slice(0, firstRepeat)
  const rows: ArgumentRow[] = fixed.map((param, index) => ({
    name: param.name,
    detail: param.detail,
    require: param.require,
    index,
  }))
  if (firstRepeat === -1) return rows
  const group = params.slice(firstRepeat)
  const filled = Math.max(0, argCount - fixed.length)
  const repetitions = Math.min(
    Math.floor((MAX_ARGUMENT_ROWS - fixed.length) / group.length),
    Math.floor(filled / group.length) + 1,
  )
  for (let k = 0; k < repetitions; k++) {
    for (const [offset, param] of group.entries()) {
      rows.push({
        name: k === 0 ? param.name : repeatedName(param.name, k + 1),
        detail: param.detail,
        require: param.require && k === 0,
        index: fixed.length + k * group.length + offset,
      })
    }
  }
  return rows
}

export type PreviewValue =
  | number
  | string
  | boolean
  | null
  | undefined
  | ReadonlyArray<ReadonlyArray<number | string | boolean | null>>

function formatScalar(value: number | string | boolean | null): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number')
    return Number.isInteger(value) ? String(value) : String(+value.toPrecision(10))
  if (/^#[A-Z/0-9!?]+$/.test(value)) return value
  return `"${value}"`
}

const PREVIEW_CELLS = 6

/// Excel's preview text: `42`, `"text"`, `TRUE`, `{1;2;3}` (rows joined by
/// `;`, columns by `,`), long arrays elided with `...`.
export function formatPreviewValue(value: PreviewValue): string {
  if (value === undefined || value === null) return ''
  if (!Array.isArray(value)) return formatScalar(value as number | string | boolean)
  const rows = value as ReadonlyArray<ReadonlyArray<number | string | boolean | null>>
  if (rows.length === 1 && rows[0]!.length === 1) return formatScalar(rows[0]![0]!)
  const cells: string[] = []
  let shown = 0
  let elided = false
  for (const row of rows) {
    const cols: string[] = []
    for (const cell of row) {
      if (shown >= PREVIEW_CELLS) {
        elided = true
        break
      }
      cols.push(formatScalar(cell))
      shown++
    }
    if (cols.length > 0) cells.push(cols.join(','))
    if (elided) break
  }
  return `{${cells.join(';')}${elided ? ';...' : ''}}`
}

/// Excel's factory list for Most Recently Used.
export const DEFAULT_RECENT_FUNCTIONS: readonly string[] = [
  'SUM',
  'AVERAGE',
  'IF',
  'HYPERLINK',
  'COUNT',
  'MAX',
  'SIN',
  'SUMIF',
  'PMT',
  'STDEV',
]

export const RECENT_FUNCTIONS_KEY = 'genoffice.sheets.recentFunctions'

/// Sheet names need quoting unless they are a bare identifier that cannot be
/// read as a cell reference.
export function quoteSheetName(name: string): string {
  const bare = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
  return bare ? name : `'${name.replace(/'/g, "''")}'`
}

export function columnLabel(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

export interface PickedRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  /// 1 = whole rows, 2 = whole columns (Univer's RANGE_TYPE).
  readonly rangeType?: number
}

/// Relative A1 text for a picked range, sheet-qualified only when the pick
/// is on another sheet than the formula's cell.
export function pickedRangeRef(
  range: PickedRange,
  sheetName: string,
  targetSheetName: string,
): string {
  let body: string
  if (range.rangeType === 2) {
    const from = columnLabel(range.startColumn)
    const to = columnLabel(range.endColumn)
    body = `${from}:${to}`
  } else if (range.rangeType === 1) {
    body = `${range.startRow + 1}:${range.endRow + 1}`
  } else {
    const from = `${columnLabel(range.startColumn)}${range.startRow + 1}`
    const to = `${columnLabel(range.endColumn)}${range.endRow + 1}`
    body = from === to ? from : `${from}:${to}`
  }
  return sheetName === targetSheetName ? body : `${quoteSheetName(sheetName)}!${body}`
}
