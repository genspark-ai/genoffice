/**
 * Cell text → typed value classification (P26). The success criterion of a
 * spreadsheet export is that numeric-looking table cells become real numbers
 * the user can compute with — while codes, phone numbers, leading-zero
 * strings and dates stay text (mirrors apps/sheets csv-import's contract;
 * the plain-number regex is kept in sync with its isNumericCell).
 */

/** builtin numFmt ids (ECMA-376 §18.8.30); custom codes get ids ≥ 164 */
export const BUILTIN_NUM_FMTS: Record<string, number> = {
  General: 0,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
}

export interface ParsedNumber {
  kind: 'number'
  value: number
  /** numFmt code; 'General' for plain numbers */
  numFmt: string
}

export interface ParsedText {
  kind: 'text'
  text: string
}

export type ParsedCell = ParsedNumber | ParsedText

/** plain decimal (csv-import's isNumericCell): leading zeros stay text */
const PLAIN_NUMBER = /^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?$/
/** comma-grouped decimal: groups of exactly three after the first 1–3 digits */
const GROUPED_NUMBER = /^-?[1-9][0-9]{0,2}(,[0-9]{3})+(\.[0-9]+)?$/
/** currency symbols accepted as a prefix (v1: $, ¥/￥, €, £); the minus may
 * sit on either side of the symbol ("-$45" and "$-45" both parse) */
const CURRENCY_PREFIX = /^(-?)([$¥￥€£])\s?(-?[0-9.,]+)$/
/** currency words/symbols accepted as a suffix (v1: CNY yuan character, euro, dollar) */
const CURRENCY_SUFFIX = /^(-?[0-9.,]+)\s?(元|€|\$)$/
/** percent: numeric part then '%' */
const PERCENT = /^(-?[0-9.,]+)\s?%$/

/**
 * integers longer than 15 significant digits lose precision as IEEE doubles
 * (serials, card numbers) — keep them text
 */
const MAX_INTEGER_DIGITS = 15

/** '¥' normalizes over the full-width '￥' (same currency, one style pool entry) */
const CURRENCY_FMT: Record<string, { int: string; dec: string }> = {
  $: { int: '"$"#,##0', dec: '"$"#,##0.00' },
  '¥': { int: '"¥"#,##0', dec: '"¥"#,##0.00' },
  '€': { int: '"€"#,##0', dec: '"€"#,##0.00' },
  '£': { int: '"£"#,##0', dec: '"£"#,##0.00' },
  元: { int: '#,##0"元"', dec: '#,##0.00"元"' },
}

/** digits after the decimal point in the source text (0 = integer) */
function decimalDigits(text: string): number {
  const dot = text.indexOf('.')
  return dot < 0 ? 0 : text.length - dot - 1
}

function significantIntegerDigits(text: string): number {
  return text.replace(/[-,.]/g, '').replace(/^0+/, '').length
}

/**
 * the bare numeric part: plain or comma-grouped decimal. Returns null for
 * anything else — leading zeros, misplaced groups, dates, exponents-with-
 * commas all fall through to text.
 */
function parseBareNumber(text: string): { value: number; grouped: boolean } | null {
  if (PLAIN_NUMBER.test(text)) {
    // >15-digit integers (phone-adjacent serials) would silently round
    if (!text.includes('.') && !/[eE]/.test(text)) {
      if (significantIntegerDigits(text) > MAX_INTEGER_DIGITS) return null
    }
    const value = Number(text)
    return Number.isFinite(value) ? { value, grouped: false } : null
  }
  if (GROUPED_NUMBER.test(text)) {
    const value = Number(text.replaceAll(',', ''))
    return Number.isFinite(value) ? { value, grouped: true } : null
  }
  return null
}

/** percent format code matching the source's decimal places, e.g. "45.3%" → "0.0%" */
function percentFmt(digits: number): string {
  return digits === 0 ? '0%' : `0.${'0'.repeat(digits)}%`
}

/**
 * Classify one cell's text. Whitespace is trimmed for detection but the
 * returned text (when it stays text) is the caller's original string —
 * classification never rewrites content.
 */
export function parseCellValue(raw: string): ParsedCell {
  const text = raw.trim()
  const asText: ParsedText = { kind: 'text', text: raw }
  if (text === '') return asText

  const bare = parseBareNumber(text)
  if (bare) {
    return {
      kind: 'number',
      value: bare.value,
      numFmt: bare.grouped ? (text.includes('.') ? '#,##0.00' : '#,##0') : 'General',
    }
  }

  const percent = PERCENT.exec(text)
  if (percent) {
    const num = parseBareNumber(percent[1]!)
    if (num) {
      // ×100 sources are exact decimals; /100 reintroduces float noise
      // (45.3 / 100 = 0.45299999…) → round to the source's precision
      const digits = decimalDigits(percent[1]!) + 2
      return {
        kind: 'number',
        value: Number((num.value / 100).toFixed(digits)),
        numFmt: percentFmt(decimalDigits(percent[1]!)),
      }
    }
    return asText
  }

  const prefix = CURRENCY_PREFIX.exec(text)
  if (prefix) {
    // a minus on both sides ("-$-45") is noise, not a number
    if (prefix[1] === '-' && prefix[3]!.startsWith('-')) return asText
    const num = parseBareNumber(prefix[3]!)
    if (num) {
      const fmt = CURRENCY_FMT[prefix[2] === '￥' ? '¥' : prefix[2]!]!
      return {
        kind: 'number',
        value: prefix[1] === '-' ? -num.value : num.value,
        numFmt: decimalDigits(prefix[3]!) > 0 ? fmt.dec : fmt.int,
      }
    }
    return asText
  }

  const suffix = CURRENCY_SUFFIX.exec(text)
  if (suffix) {
    const num = parseBareNumber(suffix[1]!)
    if (num) {
      const fmt = CURRENCY_FMT[suffix[2]!]!
      return {
        kind: 'number',
        value: num.value,
        numFmt: decimalDigits(suffix[1]!) > 0 ? fmt.dec : fmt.int,
      }
    }
    return asText
  }

  return asText
}
