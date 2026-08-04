import { numfmt } from '@univerjs/core'

/// Model behind the Format Cells → Number tab: category + sub-options ⇄
/// OOXML pattern. The ribbon's category dropdown (number-format.ts) derives
/// its representative patterns from the same builders.

export type NumfmtCategory =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'date'
  | 'time'
  | 'percentage'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'custom'

export type NegativeStyle = 'minus' | 'red' | 'parens' | 'red-parens'

export interface NumfmtOptions {
  readonly category: NumfmtCategory
  readonly decimals: number
  readonly thousands: boolean
  readonly negative: NegativeStyle
  /// Currency/accounting symbol; '' = none.
  readonly symbol: string
  readonly datePattern: string
  readonly timePattern: string
  readonly fractionPattern: string
  readonly custom: string
}

export const NUMFMT_CATEGORIES: readonly NumfmtCategory[] = [
  'general',
  'number',
  'currency',
  'accounting',
  'date',
  'time',
  'percentage',
  'fraction',
  'scientific',
  'text',
  'custom',
]

export const NEGATIVE_STYLES: readonly NegativeStyle[] = ['minus', 'red', 'parens', 'red-parens']

export const CURRENCY_SYMBOLS: readonly string[] = ['¥', '$', '€', '£', 'US$', 'HK$', '₩']

export const DATE_PATTERNS: readonly string[] = [
  'yyyy-mm-dd',
  'yyyy/m/d',
  'm/d/yyyy',
  'd-mmm-yy',
  'mmm-yy',
  'mmmm d, yyyy',
  'dddd, mmmm dd, yyyy',
  'yyyy"年"m"月"d"日"',
  'm"月"d"日"',
  'yyyy-mm-dd hh:mm',
]

export const TIME_PATTERNS: readonly string[] = [
  'h:mm',
  'h:mm:ss',
  'hh:mm:ss',
  'h:mm AM/PM',
  'h:mm:ss AM/PM',
  '[h]:mm:ss',
  'mm:ss',
]

export const FRACTION_PATTERNS: readonly string[] = [
  '# ?/?',
  '# ??/??',
  '# ???/???',
  '# ?/2',
  '# ?/4',
  '# ?/8',
  '# ??/16',
  '# ?/10',
]

export const DEFAULT_NUMFMT_OPTIONS: NumfmtOptions = {
  category: 'general',
  decimals: 2,
  thousands: false,
  negative: 'minus',
  symbol: '$',
  datePattern: 'yyyy-mm-dd',
  timePattern: 'h:mm:ss',
  fractionPattern: '# ?/?',
  custom: '',
}

export function clampDecimals(value: number): number {
  if (!Number.isFinite(value)) return 0
  return Math.min(30, Math.max(0, Math.trunc(value)))
}

function digits(decimals: number): string {
  const count = clampDecimals(decimals)
  return count > 0 ? `.${'0'.repeat(count)}` : ''
}

/// Single-char symbols stay bare ($#,##0); multi-char ones need quoting.
function symbolToken(symbol: string): string {
  if (!symbol) return ''
  return symbol.length === 1 ? symbol : `"${symbol}"`
}

/// Excel's four negative-number styles around a positive-section base.
function withNegative(base: string, negative: NegativeStyle): string {
  switch (negative) {
    case 'minus':
      return base
    case 'red':
      return `${base};[Red]${base}`
    case 'parens':
      return `${base}_);(${base})`
    case 'red-parens':
      return `${base}_);[Red](${base})`
  }
}

export function numfmtPattern(options: NumfmtOptions): string {
  switch (options.category) {
    case 'general':
      return 'General'
    case 'number':
      return withNegative(
        `${options.thousands ? '#,##0' : '0'}${digits(options.decimals)}`,
        options.negative,
      )
    case 'currency':
      return withNegative(
        `${symbolToken(options.symbol)}#,##0${digits(options.decimals)}`,
        options.negative,
      )
    case 'accounting': {
      const symbol = symbolToken(options.symbol)
      const base = `#,##0${digits(options.decimals)}`
      const zero = `"-"${'?'.repeat(clampDecimals(options.decimals))}`
      return `_(${symbol}* ${base}_);_(${symbol}* (${base});_(${symbol}* ${zero}_);_(@_)`
    }
    case 'date':
      return options.datePattern
    case 'time':
      return options.timePattern
    case 'percentage':
      return `0${digits(options.decimals)}%`
    case 'fraction':
      return options.fractionPattern
    case 'scientific':
      return `0${digits(options.decimals)}E+00`
    case 'text':
      return '@'
    case 'custom':
      return options.custom
  }
}

/// Best-effort reverse of numfmtPattern: extract candidate options, rebuild,
/// and only accept on an exact round-trip; anything else lands in Custom.
export function numfmtOptionsOf(pattern: string): NumfmtOptions {
  const defaults = DEFAULT_NUMFMT_OPTIONS
  if (!pattern || pattern === 'General') return { ...defaults, category: 'general' }
  if (pattern === '@') return { ...defaults, category: 'text' }
  if (DATE_PATTERNS.includes(pattern))
    return { ...defaults, category: 'date', datePattern: pattern }
  if (TIME_PATTERNS.includes(pattern))
    return { ...defaults, category: 'time', timePattern: pattern }
  if (FRACTION_PATTERNS.includes(pattern)) {
    return { ...defaults, category: 'fraction', fractionPattern: pattern }
  }
  let match = /^0(?:\.(0+))?%$/.exec(pattern)
  if (match) return { ...defaults, category: 'percentage', decimals: match[1]?.length ?? 0 }
  match = /^0(?:\.(0+))?E\+00$/.exec(pattern)
  if (match) return { ...defaults, category: 'scientific', decimals: match[1]?.length ?? 0 }
  match = /^_\((?:"([^"]+)"|([^*\s]))?\* #,##0(?:\.(0+))?_\)/.exec(pattern)
  if (match) {
    const candidate: NumfmtOptions = {
      ...defaults,
      category: 'accounting',
      symbol: match[1] ?? match[2] ?? '',
      decimals: match[3]?.length ?? 0,
    }
    if (numfmtPattern(candidate) === pattern) return candidate
  }
  match = /^(?:"([^"]+)"|([^0#".,;[\]_]))?(#,##0|0)(?:\.(0+))?(?:[_;].*)?$/.exec(pattern)
  if (match) {
    const symbol = match[1] ?? match[2] ?? ''
    for (const negative of NEGATIVE_STYLES) {
      const candidate: NumfmtOptions = symbol
        ? { ...defaults, category: 'currency', symbol, decimals: match[4]?.length ?? 0, negative }
        : {
            ...defaults,
            category: 'number',
            thousands: match[3] === '#,##0',
            decimals: match[4]?.length ?? 0,
            negative,
          }
      if (numfmtPattern(candidate) === pattern) return candidate
    }
  }
  return { ...defaults, category: 'custom', custom: pattern }
}

export function numfmtPreview(pattern: string, value: number | string): string {
  try {
    return numfmt.format(pattern || 'General', value, { throws: false })
  } catch {
    return typeof value === 'string' ? value : String(value)
  }
}

/// Excel 1900-system serial for "now", so date/time previews show today.
export function todaySerial(now = new Date()): number {
  const utc = Date.UTC(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    now.getHours(),
    now.getMinutes(),
    now.getSeconds(),
  )
  return utc / 86400000 + 25569
}

/// Preview value: the anchor cell's number when there is one, else a
/// category-appropriate stand-in.
export function sampleValue(
  category: NumfmtCategory,
  anchor: number | string | null,
  serial: number,
): number | string {
  if (typeof anchor === 'number') return anchor
  if (category === 'date' || category === 'time') return serial
  if (typeof anchor === 'string' && anchor !== '') return anchor
  return 1234.56
}
