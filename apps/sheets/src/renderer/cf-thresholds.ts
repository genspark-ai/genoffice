/**
 * Scale thresholds (cfvo) for colorScale / dataBar / iconSet conditional
 * formats, resolved against the sidecar's cached cell values, plus the
 * dataBar geometry Excel derives from attributes Univer's bar painter does
 * not model.
 *
 * Why not Univer's formula engine: the renderer streams a window of cells
 * into Univer and evicts the rest, so a `CFValueType.formula` threshold
 * registered with the engine evaluates against whatever happens to be
 * loaded (blank off-window cells read as 0), and `Sheet2!$A$1` on a sheet
 * that never streamed is always 0. The sidecar cache holds the file's real
 * values, so thresholds are resolved here once per rule and installed as
 * static `num` stops.
 *
 * Excel semantics honoured here (verified on the LibreOffice corpus refs):
 * relative references inside a scale threshold read as 0 (colorscale.xlsx
 * F3:F6 `2*A1+2` renders as 2 with A1=1); `$`-absolute references read the
 * cell. Formulas the evaluator cannot fold fall back to the rule's default
 * threshold for that slot (lowest / highest value, 50th percentile, or the
 * icon set's even percent split) — never to a stop Univer cannot compute,
 * which would leave the rule unpainted.
 */
import { parseAddress, type RangeBounds } from '../domain/cell-address'

export interface ScaleCfvo {
  kind: string
  value?: string | undefined
  gte?: boolean | undefined
}

/// Sidecar-backed lookups the evaluator needs. `sheetName === null` means
/// the rule's own sheet.
export interface ThresholdReader {
  /// Numeric values of the populated cells of `range` (NaN for text cells,
  /// blanks omitted), or null when the sheet is unknown, the range is too
  /// large, or the read failed.
  readValues(sheetName: string | null, range: RangeBounds): Promise<number[] | null>
  /// Formula text of a workbook-defined name (case-insensitive), or null.
  definedName(name: string): string | null
  /// Sheet and body range of a table column for `Table[Column]`, or null.
  tableColumn(table: string, column: string): { sheetName: string; range: RangeBounds } | null
}

/// Above this many cells an aggregate threshold falls back to the default
/// stop instead of paying a chunked full-range read on open (same ceiling as
/// the auto-bounds resolution).
export const THRESHOLD_RANGE_CELL_CAP = 512_000

/// Nested defined names resolve this many levels deep.
const NAME_DEPTH_LIMIT = 3

// Cell reference with an optional sheet prefix; the second cell of a range
// never carries a prefix.
const SHEET_PREFIX = String.raw`(?:(?:'([^']+)'|([A-Za-z0-9_.]+))!)?`
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})`
const RANGE_ARGUMENT = new RegExp(
  String.raw`^\s*${SHEET_PREFIX}${CELL}(?::${CELL})?\s*(?:,\s*(-?[0-9]+(?:\.[0-9]+)?)\s*)?$`,
)
const AGGREGATE_CALL =
  /(?<![\w.])(AVERAGE|MIN|MAX|SUM|COUNT|MEDIAN|PERCENTILE(?:\.INC)?|QUARTILE(?:\.INC)?)\(([^()]*)\)/gi
const STRUCTURED_CALL =
  /(?<![\w.])(AVERAGE|MIN|MAX|SUM|COUNT|MEDIAN)\(\s*([A-Za-z_][\w.]*)\[([^\]]+)\]\s*\)/gi
const CELL_REFERENCE = new RegExp(String.raw`(?<![\w$.!'])${SHEET_PREFIX}${CELL}(?![\w(])`, 'g')
const NAME_TOKEN = /(?<![\w$.!'])[A-Za-z_\\][\w.]*(?![\w.(![])/g

/// Folds a threshold formula body (without the leading `=`) to a number:
/// `$T$5`, `Sheet2!$B$2`, `AVERAGE($B$2:$B$20)*1.1`, `PERCENTILE($A$1:$A$9,0.9)`,
/// `SUM(Table[Amount])`, defined names, and +-*/() arithmetic over those.
/// Relative references read as 0. Returns null when the body needs
/// anything else (string literals, other functions, whole-column ranges).
export async function evaluateThresholdFormula(
  body: string,
  reader: ThresholdReader,
  depth = 0,
): Promise<number | null> {
  let expression = body.replace(/^=/, '').trim()
  if (expression === '' || expression.includes('"') || depth > NAME_DEPTH_LIMIT) return null

  const structured = await replaceAsync(expression, STRUCTURED_CALL, async (match) => {
    const column = reader.tableColumn(match[2]!, match[3]!)
    if (!column) return null
    const values = await reader.readValues(column.sheetName, column.range)
    if (values === null) return null
    return aggregate(match[1]!.toUpperCase(), values.filter(Number.isFinite), undefined)
  })
  if (structured === null) return null
  expression = structured

  const aggregated = await replaceAsync(expression, AGGREGATE_CALL, async (match) => {
    const argument = RANGE_ARGUMENT.exec(match[2]!)
    if (!argument) return null
    const [, quoted, bare, c1, col1, r1, row1, c2, col2, r2, row2, k] = argument
    // A relative reference anywhere in the range reads as 0, so the
    // aggregate itself folds to 0 (Excel's scale-threshold semantics).
    if (c1 !== '$' || r1 !== '$' || (col2 !== undefined && (c2 !== '$' || r2 !== '$'))) {
      return 0
    }
    const first = parseAddress(`${col1!.toUpperCase()}${row1}`)
    const second = col2 === undefined ? first : parseAddress(`${col2.toUpperCase()}${row2}`)
    const range: RangeBounds = {
      startRow: Math.min(first.row, second.row),
      endRow: Math.max(first.row, second.row),
      startColumn: Math.min(first.column, second.column),
      endColumn: Math.max(first.column, second.column),
    }
    const values = await reader.readValues(quoted ?? bare ?? null, range)
    if (values === null) return null
    return aggregate(
      match[1]!.toUpperCase(),
      values.filter(Number.isFinite),
      k === undefined ? undefined : Number(k),
    )
  })
  if (aggregated === null) return null
  expression = aggregated

  const referenced = await replaceAsync(expression, CELL_REFERENCE, async (match) => {
    const [, quoted, bare, columnAnchor, column, rowAnchor, row] = match
    if (columnAnchor !== '$' || rowAnchor !== '$') return 0
    const cell = parseAddress(`${column!.toUpperCase()}${row}`)
    const values = await reader.readValues(quoted ?? bare ?? null, {
      startRow: cell.row,
      endRow: cell.row,
      startColumn: cell.column,
      endColumn: cell.column,
    })
    if (values === null) return null
    // A blank cell is 0 in arithmetic; a text cell is #VALUE!.
    if (values.length === 0) return 0
    return Number.isFinite(values[0]) ? values[0]! : null
  })
  if (referenced === null) return null
  expression = referenced

  const named = await replaceAsync(expression, NAME_TOKEN, async (match) => {
    const formula = reader.definedName(match[0])
    if (formula === null) return null
    return evaluateThresholdFormula(formula, reader, depth + 1)
  })
  if (named === null) return null
  return evaluateArithmetic(named)
}

/// True when the formula reads no cell, name, or table — only literals and
/// function calls (`TODAY()-30`). Univer's formula engine evaluates such a
/// threshold correctly whatever the streamed window holds, so an unfoldable
/// self-contained formula may stay a formula stop instead of taking the
/// slot default.
export function isSelfContainedFormula(body: string): boolean {
  const bare = body.replace(/^=/, '').replace(/"[^"]*"/g, '""')
  if (/[$![']/.test(bare)) return false
  if (/(?<![\w.])[A-Za-z]{1,3}[0-9]{1,7}(?![\w(])/.test(bare)) return false
  // Every remaining identifier must be a function call.
  return [...bare.matchAll(/(?<![\w.])[A-Za-z_][\w.]*/g)].every((match) =>
    /^\s*\(/.test(bare.slice(match.index + match[0].length)),
  )
}

/// Sequential regex replacement with an async resolver; a null resolution
/// aborts the whole fold.
async function replaceAsync(
  input: string,
  pattern: RegExp,
  resolve: (match: RegExpExecArray) => Promise<number | null>,
): Promise<string | null> {
  const matches = [...input.matchAll(pattern)] as RegExpExecArray[]
  if (matches.length === 0) return input
  let output = ''
  let cursor = 0
  for (const match of matches) {
    const value = await resolve(match)
    if (value === null || !Number.isFinite(value)) return null
    output += input.slice(cursor, match.index) + numberLiteral(value)
    cursor = match.index + match[0].length
  }
  return output + input.slice(cursor)
}

/// Parenthesised so a negative substitution never fuses with the operator
/// before it (`3-(-5)`); the arithmetic parser reads exponent notation.
function numberLiteral(value: number): string {
  return `(${String(value)})`
}

/// Excel's PERCENTILE.INC / MEDIAN / QUARTILE.INC over ascending `sorted`.
export function percentileInc(sorted: readonly number[], k: number): number | null {
  if (sorted.length === 0 || !Number.isFinite(k) || k < 0 || k > 1) return null
  const rank = (sorted.length - 1) * k
  const lower = Math.floor(rank)
  const upper = Math.min(lower + 1, sorted.length - 1)
  return sorted[lower]! + (sorted[upper]! - sorted[lower]!) * (rank - lower)
}

function aggregate(name: string, values: number[], k: number | undefined): number | null {
  switch (name) {
    case 'SUM':
      return values.reduce((sum, value) => sum + value, 0)
    case 'COUNT':
      return values.length
    // MIN / MAX of an all-blank range are 0 in Excel. Looped rather than
    // spread: a range at THRESHOLD_RANGE_CELL_CAP exceeds the engine's
    // call-argument limit and `Math.min(...values)` would throw.
    case 'MIN': {
      let low = Number.POSITIVE_INFINITY
      for (const value of values) if (value < low) low = value
      return values.length === 0 ? 0 : low
    }
    case 'MAX': {
      let high = Number.NEGATIVE_INFINITY
      for (const value of values) if (value > high) high = value
      return values.length === 0 ? 0 : high
    }
    case 'AVERAGE':
      return values.length === 0
        ? null
        : values.reduce((sum, value) => sum + value, 0) / values.length
    case 'MEDIAN':
      return percentileInc(
        [...values].sort((a, b) => a - b),
        0.5,
      )
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
      return k === undefined
        ? null
        : percentileInc(
            [...values].sort((a, b) => a - b),
            k,
          )
    case 'QUARTILE':
    case 'QUARTILE.INC':
      return k === undefined || !Number.isInteger(k)
        ? null
        : percentileInc(
            [...values].sort((a, b) => a - b),
            k / 4,
          )
    default:
      return null
  }
}

/// Tiny +-*/() evaluator so a substituted threshold ("2*(0)+3") becomes a
/// static num stop the monotonic clamp can see; anything else returns null.
export function evaluateArithmetic(expression: string): number | null {
  const source = expression.replace(/^=/, '').replace(/\s+/g, '')
  if (source === '' || !/^[\d+\-*/().eE]+$/.test(source)) return null
  let position = 0
  const parseExpression = (): number => {
    let value = parseTerm()
    while (source[position] === '+' || source[position] === '-') {
      const operator = source[position]
      position += 1
      const term = parseTerm()
      value = operator === '+' ? value + term : value - term
    }
    return value
  }
  const parseTerm = (): number => {
    let value = parseFactor()
    while (source[position] === '*' || source[position] === '/') {
      const operator = source[position]
      position += 1
      const factor = parseFactor()
      value = operator === '*' ? value * factor : value / factor
    }
    return value
  }
  const parseFactor = (): number => {
    if (source[position] === '-') {
      position += 1
      return -parseFactor()
    }
    if (source[position] === '+') {
      position += 1
      return parseFactor()
    }
    if (source[position] === '(') {
      position += 1
      const value = parseExpression()
      if (source[position] !== ')') return Number.NaN
      position += 1
      return value
    }
    const match = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(source.slice(position))
    if (!match) return Number.NaN
    position += match[0].length
    return Number(match[0])
  }
  const value = parseExpression()
  return position === source.length && Number.isFinite(value) ? value : null
}

/// The threshold Excel's dialog would give the slot when the file's formula
/// cannot be folded: lowest / highest value at the ends, the median for a
/// three-color midpoint, an even percent split for icon sets.
export function defaultThreshold(ruleType: string, index: number, count: number): ScaleCfvo {
  if (index <= 0) return { kind: 'min' }
  if (ruleType === 'iconSet') {
    return { kind: 'percent', value: String(Math.round((index * 100) / count)) }
  }
  if (index >= count - 1) return { kind: 'max' }
  return { kind: 'percentile', value: '50' }
}

/// Excel forces color-scale thresholds to be non-decreasing: a later stop
/// below an earlier one snaps up to it, so values past the earlier stop take
/// the last color solid. Univer interpolates the stops verbatim, so replicate
/// the clamp; equal neighbors get an epsilon step downward so the later color
/// wins at the shared boundary, matching Excel.
export function clampColorScaleStops<T extends { kind: string; value?: string | undefined }>(
  cfvos: T[],
): T[] {
  const stops = cfvos.map((cfvo) =>
    cfvo.kind === 'num' && cfvo.value !== undefined && Number.isFinite(Number(cfvo.value))
      ? Number(cfvo.value)
      : null,
  )
  let previous: number | null = null
  const clamped = stops.map((stop) => {
    if (stop === null) {
      previous = null
      return null
    }
    const lifted = previous !== null && stop < previous ? previous : stop
    previous = lifted
    return lifted
  })
  for (let i = clamped.length - 1; i > 0; i -= 1) {
    const current = clamped[i] ?? null
    const before = clamped[i - 1] ?? null
    if (current !== null && before !== null && before >= current) {
      clamped[i - 1] = current - Math.max(Math.abs(current) * 1e-9, 1e-9)
    }
  }
  if (clamped.every((stop, index) => stop === null || stop === stops[index])) return cfvos
  return cfvos.map((cfvo, index) => {
    const stop = clamped[index] ?? null
    return stop === null || stop === stops[index] ? cfvo : { ...cfvo, value: String(stop) }
  })
}

// ---------------------------------------------------------------------------
// dataBar geometry
// ---------------------------------------------------------------------------

/// What the renderer knows about a dataBar rule after its thresholds have
/// been folded to numbers where possible.
export interface DataBarLayoutInput {
  readonly cfvos: readonly ScaleCfvo[]
  /// Effective extents from the sidecar (% of the cell width); undefined
  /// from an older sidecar means Univer's native 0..100.
  readonly minLength?: number | undefined
  readonly maxLength?: number | undefined
  readonly axisPosition?: string | undefined
}

const DATA_DRIVEN_KINDS = new Set(['min', 'max', 'autoMin', 'autoMax', 'percent', 'percentile'])

/// Whether laying the bar out needs the rule's cached cell values.
export function dataBarNeedsValues(input: DataBarLayoutInput): boolean {
  if (!dataBarNeedsLayout(input)) return false
  return input.cfvos.some((cfvo) => DATA_DRIVEN_KINDS.has(cfvo.kind))
}

/// True when Univer's native min/max mapping would not match Excel: x14
/// autoMin/autoMax bounds, non-0/100 extents, or a cell-midpoint axis.
export function dataBarNeedsLayout(input: DataBarLayoutInput): boolean {
  const [low, high] = input.cfvos
  if (!low || !high) return false
  if (low.kind === 'autoMin' || low.kind === 'autoMax') return true
  if (high.kind === 'autoMin' || high.kind === 'autoMax') return true
  const minLength = input.minLength ?? 0
  const maxLength = input.maxLength ?? 100
  return minLength !== 0 || maxLength !== 100 || input.axisPosition === 'middle'
}

/// A threshold as Excel resolves it for a data bar over `values` (the rule
/// range's numeric cells): the file's number, the data extremes, the
/// zero-anchored auto extremes, or a percent / percentile of the data.
export function resolveBarBound(
  cfvo: ScaleCfvo,
  values: readonly number[] | null,
  side: 'min' | 'max',
): number | null {
  if (cfvo.kind === 'num') {
    const value = Number(cfvo.value)
    return Number.isFinite(value) ? value : null
  }
  if (values === null) return null
  const finite = values.filter(Number.isFinite)
  if (finite.length === 0) return null
  let low = Number.POSITIVE_INFINITY
  let high = Number.NEGATIVE_INFINITY
  for (const value of finite) {
    if (value < low) low = value
    if (value > high) high = value
  }
  switch (cfvo.kind) {
    case 'min':
      return low
    case 'max':
      return high
    // Excel anchors auto bounds at zero for one-signed data.
    case 'autoMin':
      return Math.min(0, low)
    case 'autoMax':
      return Math.max(0, high)
    case 'percent': {
      const percent = Number(cfvo.value ?? (side === 'min' ? 0 : 100))
      if (!Number.isFinite(percent)) return null
      return low + (Math.max(0, Math.min(100, percent)) / 100) * (high - low)
    }
    case 'percentile': {
      const percent = Number(cfvo.value ?? (side === 'min' ? 0 : 100))
      if (!Number.isFinite(percent)) return null
      return percentileInc(
        [...finite].sort((a, b) => a - b),
        Math.max(0, Math.min(100, percent)) / 100,
      )
    }
    default:
      return null
  }
}

/// Excel maps [min, max] onto [minLength%, maxLength%] of the cell width; a
/// legacy (2006-only) bar draws its smallest value 10% wide and its largest
/// 90% wide. Univer maps its bounds onto 0..100%, so widen them: the
/// returned bounds make Univer's linear bar coincide with Excel's for every
/// value in [min, max]. Null when nothing to do, or when the widened lower
/// bound would cross zero — Univer then switches to its two-sided axis
/// layout (bars start at the axis, not the cell edge), which looks worse
/// than the plain 0..100 mapping, so the caller keeps Univer's native bar.
export function emulateBarExtents(
  min: number,
  max: number,
  minLength: number,
  maxLength: number,
): { min: number; max: number } | null {
  if (minLength === 0 && maxLength === 100) return null
  if (!(max > min) || !(maxLength > minLength) || minLength < 0 || maxLength > 100) return null
  const a = minLength / 100
  const b = maxLength / 100
  const span = (max - min) / (b - a)
  const lowered = min - a * span
  if (lowered < 0) return null
  return { min: lowered, max: lowered + span }
}

/// x14 axisPosition="middle": the axis sits at the cell midpoint and both
/// halves share one scale (the larger absolute bound), so a bar's length
/// stays proportional to its magnitude. Univer derives the axis from the
/// bounds' signs, so symmetric bounds put it in the middle. Only for
/// two-signed data — Excel draws one-signed data from the cell edge.
export function middleAxisBounds(min: number, max: number): { min: number; max: number } | null {
  if (!(min < 0 && max > 0)) return null
  const extent = Math.max(-min, max)
  return { min: -extent, max: extent }
}

/// The cfvos to install for a dataBar rule, or null when the file's own
/// cfvos already render as Excel does. `values` are the rule range's cached
/// numeric cells (null when unavailable or over the cap).
export function layoutDataBar(
  input: DataBarLayoutInput,
  values: readonly number[] | null,
): ScaleCfvo[] | null {
  if (!dataBarNeedsLayout(input)) return null
  const [low, high] = input.cfvos as [ScaleCfvo, ScaleCfvo]
  const lower = resolveBarBound(low, values, 'min')
  const upper = resolveBarBound(high, values, 'max')
  const num = (value: number): ScaleCfvo => ({ kind: 'num', value: String(value) })
  // Auto bounds resolve independently: one side may still be unavailable.
  const autoResolved = [
    low.kind === 'autoMin' || low.kind === 'autoMax' ? (lower === null ? low : num(lower)) : low,
    high.kind === 'autoMin' || high.kind === 'autoMax'
      ? upper === null
        ? high
        : num(upper)
      : high,
  ]
  const fallback = autoResolved[0] === low && autoResolved[1] === high ? null : autoResolved
  if (lower === null || upper === null) return fallback
  if (input.axisPosition === 'middle') {
    const bounds = middleAxisBounds(lower, upper)
    if (bounds) return [num(bounds.min), num(bounds.max)]
    return fallback
  }
  const bounds = emulateBarExtents(lower, upper, input.minLength ?? 0, input.maxLength ?? 100)
  if (bounds) return [num(bounds.min), num(bounds.max)]
  return fallback
}
