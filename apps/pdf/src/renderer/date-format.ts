/** Acrobat AFDate_Keystroke(n) index table (the legacy non-Ex variant) */
const LEGACY_FORMATS = [
  'm/d',
  'm/d/yy',
  'mm/dd/yy',
  'mm/yy',
  'd-mmm',
  'd-mmm-yy',
  'dd-mmm-yy',
  'yy-mm-dd',
  'mmm-yy',
  'mmmm-yy',
  'mmm d, yyyy',
  'mmmm d, yyyy',
  'm/d/yy h:MM tt',
  'm/d/yy HH:MM',
]

/** Date format an AcroForm text field validates with, read from its pdf.js `actions` (AA /K and /F scripts). */
export function dateFormatFromActions(actions: unknown): string | null {
  if (!actions || typeof actions !== 'object') return null
  const { Keystroke, Format } = actions as { Keystroke?: unknown; Format?: unknown }
  const scripts = [
    ...(Array.isArray(Keystroke) ? Keystroke : []),
    ...(Array.isArray(Format) ? Format : []),
  ]
  for (const script of scripts) {
    if (typeof script !== 'string') continue
    const ex = /AFDate_(?:Keystroke|Format)Ex\s*\(\s*(["'])(.*?)\1/.exec(script)
    if (ex?.[2]) return ex[2]
    const legacy = /AFDate_(?:Keystroke|Format)\s*\(\s*(\d+)\s*\)/.exec(script)
    if (legacy) return LEGACY_FORMATS[Number(legacy[1])] ?? null
  }
  return null
}

const MONTHS = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
]

type Part = 'year' | 'month' | 'day' | 'hour' | 'hour12' | 'minute' | 'second' | 'ampm'

const TOKENS: Record<string, [Part, string]> = {
  yyyy: ['year', '\\d{4}'],
  yy: ['year', '\\d{2}'],
  mmmm: ['month', '[A-Za-z]{3,9}'],
  mmm: ['month', '[A-Za-z]{3}'],
  mm: ['month', '\\d{2}'],
  m: ['month', '\\d{1,2}'],
  dd: ['day', '\\d{2}'],
  d: ['day', '\\d{1,2}'],
  HH: ['hour', '\\d{2}'],
  H: ['hour', '\\d{1,2}'],
  hh: ['hour12', '\\d{2}'],
  h: ['hour12', '\\d{1,2}'],
  MM: ['minute', '\\d{2}'],
  M: ['minute', '\\d{1,2}'],
  ss: ['second', '\\d{2}'],
  s: ['second', '\\d{1,2}'],
  tt: ['ampm', '[AaPp][Mm]'],
  t: ['ampm', '[AaPp]'],
}
const TOKEN_RE = /yyyy|yy|mmmm|mmm|mm|m|dd|d|HH|H|hh|h|MM|M|ss|s|tt|t/g
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

type DateParts = Partial<Record<Part, number>>

/** Parse `value` written exactly in the AFDate `format`; null when it is not a real calendar date. */
function parseExact(text: string, format: string): DateParts | null {
  const parts: Part[] = []
  let re = ''
  let last = 0
  for (const m of format.matchAll(TOKEN_RE)) {
    const [part, group] = TOKENS[m[0]]!
    re += escapeRe(format.slice(last, m.index)) + `(${group})`
    parts.push(part)
    last = m.index + m[0].length
  }
  re += escapeRe(format.slice(last))
  const match = new RegExp(`^${re}$`).exec(text)
  if (!match) return null
  const got: DateParts = {}
  for (const [i, part] of parts.entries()) {
    const raw = match[i + 1]!
    if (part === 'ampm') continue
    if (part === 'month' && /[A-Za-z]/.test(raw)) {
      const month = monthFromName(raw)
      if (month === null) return null
      got.month = month
      continue
    }
    const n = Number(raw)
    got[part] = part === 'year' && raw.length === 2 ? expandYear(n) : n
  }
  return isRealDate(got) ? got : null
}

const monthFromName = (raw: string): number | null => {
  const name = raw.toLowerCase()
  const index = MONTHS.findIndex((full) => full === name || full.slice(0, 3) === name)
  return index < 0 ? null : index + 1
}
const expandYear = (n: number) => (n < 50 ? 2000 + n : 1900 + n)

function isRealDate({ year, month, day, hour, hour12, minute, second }: DateParts): boolean {
  if (month !== undefined && (month < 1 || month > 12)) return false
  if (day !== undefined) {
    const max = month !== undefined ? new Date(year ?? 2000, month, 0).getDate() : 31
    if (day < 1 || day > max) return false
  }
  if (hour !== undefined && hour > 23) return false
  if (hour12 !== undefined && (hour12 < 1 || hour12 > 12)) return false
  if (minute !== undefined && minute > 59) return false
  if (second !== undefined && second > 59) return false
  return true
}

/** Whether `value` is a real calendar date written in the AFDate `format` (empty is allowed). */
export function isValidDateInput(value: string, format: string): boolean {
  const text = value.trim()
  return !text || parseExact(text, format) !== null
}

const formatParts = (format: string): Part[] =>
  [...format.matchAll(TOKEN_RE)].map((m) => TOKENS[m[0]]![0])

/** A plain calendar date (year, month and day, no time): what the picker and loose reading can write. */
export function isDateOnlyFormat(format: string): boolean {
  const parts = formatParts(format)
  return (
    parts.length === 3 &&
    new Set(parts).size === 3 &&
    !parts.some((p) => !['year', 'month', 'day'].includes(p))
  )
}

/**
 * Year/month/day of a date typed either in `format` or in another common layout (other
 * separators, the other day/month order when unambiguous, a 4-digit year anywhere, month
 * names, or 6/8 bare digits). Only date-only formats with all three parts are read loosely.
 */
export function readDateParts(
  value: string,
  format: string,
): { year: number; month: number; day: number } | null {
  const text = value.trim()
  const exact = parseExact(text, format)
  const complete = (p: DateParts | null) =>
    p && p.year !== undefined && p.month !== undefined && p.day !== undefined
      ? { year: p.year, month: p.month, day: p.day }
      : null
  if (exact) return complete(exact)
  if (!isDateOnlyFormat(format)) return null
  let dateOrder = formatParts(format) as ('year' | 'month' | 'day')[]
  let tokens = text.split(/[\s/.,-]+/).filter(Boolean)
  if (tokens.length === 1 && /^\d{6}$|^\d{8}$/.test(tokens[0]!)) {
    const digits = tokens[0]!
    const yearLen = digits.length - 4
    let at = 0
    tokens = dateOrder.map((p) => {
      const len = p === 'year' ? yearLen : 2
      const piece = digits.slice(at, at + len)
      at += len
      return piece
    })
  }
  if (tokens.length !== 3) return null
  const got: DateParts = {}
  const rest: string[] = []
  for (const [i, tok] of tokens.entries()) {
    if (/^\d{4}$/.test(tok) && got.year === undefined) {
      got.year = Number(tok)
      // A leading 4-digit year is the ISO layout: what follows is month, then day
      if (i === 0) dateOrder = ['year', 'month', 'day']
    } else if (/^[A-Za-z]+$/.test(tok) && got.month === undefined) {
      const month = monthFromName(tok)
      if (month === null) return null
      got.month = month
    } else if (/^\d{1,2}$/.test(tok)) rest.push(tok)
    else return null
  }
  const open = dateOrder.filter((p) => got[p] === undefined)
  if (open.length !== rest.length) return null
  open.forEach((p, i) => {
    const n = Number(rest[i])
    got[p] = p === 'year' ? expandYear(n) : n
  })
  if (!isRealDate(got) && open.includes('month') && open.includes('day')) {
    ;[got.month, got.day] = [got.day, got.month]
    if (!isRealDate(got)) return null
  } else if (!isRealDate(got)) return null
  return complete(got)
}

/** Write a calendar date in the AFDate `format` (time tokens are left as typed). */
export function formatDate(
  date: { year: number; month: number; day: number },
  format: string,
): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  const monthName = MONTHS[date.month - 1]!
  return format.replace(TOKEN_RE, (tok) => {
    switch (tok) {
      case 'yyyy':
        return String(date.year)
      case 'yy':
        return pad(date.year % 100)
      case 'mmmm':
        return monthName[0]!.toUpperCase() + monthName.slice(1)
      case 'mmm':
        return monthName[0]!.toUpperCase() + monthName.slice(1, 3)
      case 'mm':
        return pad(date.month)
      case 'm':
        return String(date.month)
      case 'dd':
        return pad(date.day)
      case 'd':
        return String(date.day)
      default:
        return tok
    }
  })
}

/**
 * The value to commit for what was typed: as-is when it already matches `format`, rewritten
 * in `format` when it is a recognisable date in another layout, null when it is no date.
 */
export function normalizeDateInput(value: string, format: string): string | null {
  const text = value.trim()
  if (!text) return ''
  if (parseExact(text, format)) return text
  const parts = readDateParts(text, format)
  return parts ? formatDate(parts, format) : null
}

export const toIsoDate = (d: { year: number; month: number; day: number }): string =>
  `${String(d.year).padStart(4, '0')}-${String(d.month).padStart(2, '0')}-${String(d.day).padStart(2, '0')}`
