/** what a field's result is, which decides the switches that can shape it */
export type FieldResultKind = 'number' | 'date' | 'text'

export const fieldNameOf = (code: string): string =>
  code.trim().split(/\s+/)[0]?.toUpperCase() ?? ''

const RESULT_KINDS: Readonly<Record<string, FieldResultKind>> = {
  PAGE: 'number',
  NUMPAGES: 'number',
  DATE: 'date',
  TIME: 'date',
  CREATEDATE: 'date',
  SAVEDATE: 'date',
  FILENAME: 'text',
  MERGEFIELD: 'text',
}

/** null for a field whose result is not worked out here: a switch on it would change nothing */
export const fieldResultKind = (code: string): FieldResultKind | null =>
  RESULT_KINDS[fieldNameOf(code)] ?? null

export const NUMBER_FORMAT_SWITCHES = [
  'Arabic',
  'alphabetic',
  'ALPHABETIC',
  'roman',
  'ROMAN',
  'Ordinal',
  'CardText',
  'OrdText',
  'Hex',
  'DollarText',
] as const

export const TEXT_FORMAT_SWITCHES = ['Upper', 'Lower', 'FirstCap', 'Caps'] as const

export const NUMERIC_PICTURES = [
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '$#,##0.00;($#,##0.00)',
  '0%',
  '0.00%',
] as const

const quoted = (arg: string): string => (/[\s;]/.test(arg) ? `"${arg}"` : arg)

export const formatSwitch = (name: string): string => `\\* ${name}`
export const numericSwitch = (picture: string): string => `\\# ${quoted(picture)}`
export const dateSwitch = (picture: string): string => `\\@ "${picture}"`

const SWITCH_RE = /\\([*#@])\s*(?:"([^"]*)"|(\S+))/g

/** the `\*` format a code names, leaving out the two that only keep formatting */
export function formatSwitchOf(code: string): string | null {
  let found: string | null = null
  for (const m of code.matchAll(SWITCH_RE)) {
    const arg = m[2] ?? m[3] ?? ''
    if (m[1] === '*' && !/^(?:MERGEFORMAT|CHARFORMAT|DBCHAR)$/i.test(arg)) found = arg
  }
  return found
}

export function numericPictureOf(code: string): string | null {
  for (const m of code.matchAll(SWITCH_RE)) if (m[1] === '#') return m[2] ?? m[3] ?? ''
  return null
}

/** the code with `added` appended; a switch of the same kind already there is replaced */
export function addSwitch(code: string, added: string): string {
  const kind = added[1]
  const keep = (m: string, k: string, q: string | undefined, bare: string | undefined) =>
    k !== kind || (k === '*' && /^(?:MERGEFORMAT|CHARFORMAT|DBCHAR)$/i.test(q ?? bare ?? ''))
      ? m
      : ''
  const base = /^\\[*#@]/.test(added) ? code.replace(SWITCH_RE, keep) : code
  const tidy = base.replace(/\s{2,}/g, ' ').trim()
  // the switch that keeps formatting stays last, where Word writes it
  const tail = /\s*\\\*\s*MERGEFORMAT\b/i.exec(tidy)
  if (!tail) return `${tidy} ${added}`
  return `${tidy.slice(0, tail.index)} ${added}${tidy.slice(tail.index)}`.trim()
}

const ONES = [
  'zero',
  'one',
  'two',
  'three',
  'four',
  'five',
  'six',
  'seven',
  'eight',
  'nine',
  'ten',
  'eleven',
  'twelve',
  'thirteen',
  'fourteen',
  'fifteen',
  'sixteen',
  'seventeen',
  'eighteen',
  'nineteen',
]
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety']
const ORDINAL_WORDS: Readonly<Record<string, string>> = {
  one: 'first',
  two: 'second',
  three: 'third',
  five: 'fifth',
  eight: 'eighth',
  nine: 'ninth',
  twelve: 'twelfth',
}

function cardinalText(n: number): string {
  if (n < 20) return ONES[n]
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : '')
  if (n < 1000) {
    const rest = n % 100
    return `${ONES[Math.floor(n / 100)]} hundred${rest ? ` ${cardinalText(rest)}` : ''}`
  }
  const rest = n % 1000
  return `${cardinalText(Math.floor(n / 1000))} thousand${rest ? ` ${cardinalText(rest)}` : ''}`
}

function ordinalText(n: number): string {
  const words = cardinalText(n)
  const cut = Math.max(words.lastIndexOf(' '), words.lastIndexOf('-')) + 1
  const last = words.slice(cut)
  const ord = ORDINAL_WORDS[last] ?? (last.endsWith('y') ? `${last.slice(0, -1)}ieth` : `${last}th`)
  return words.slice(0, cut) + ord
}

function ordinalSuffix(n: number): string {
  const tens = n % 100
  if (tens >= 11 && tens <= 13) return 'th'
  return ['th', 'st', 'nd', 'rd'][n % 10] ?? 'th'
}

function roman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ]
  let out = ''
  let rest = n
  for (const [value, glyph] of table) {
    while (rest >= value) {
      out += glyph
      rest -= value
    }
  }
  return out
}

/** Word's letter numbering: a…z, then aa…zz, aaa… (the letter repeats) */
function alphabetic(n: number): string {
  return String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1)
}

function formatNumberAs(n: number, format: string): string | null {
  const whole = Math.trunc(Math.abs(n))
  const inWords = whole < 1_000_000
  switch (format) {
    case 'alphabetic':
      return whole >= 1 ? alphabetic(whole) : null
    case 'ALPHABETIC':
      return whole >= 1 ? alphabetic(whole).toUpperCase() : null
    case 'roman':
      return whole >= 1 && whole < 4000 ? roman(whole) : null
    case 'ROMAN':
      return whole >= 1 && whole < 4000 ? roman(whole).toUpperCase() : null
  }
  switch (format.toLowerCase()) {
    case 'arabic':
      return String(n)
    case 'ordinal':
      return `${whole}${ordinalSuffix(whole)}`
    case 'cardtext':
      return inWords ? cardinalText(whole) : null
    case 'ordtext':
      return inWords ? ordinalText(whole) : null
    case 'hex':
      return whole.toString(16).toUpperCase()
    case 'dollartext': {
      if (!inWords) return null
      const cents = Math.round((Math.abs(n) - whole) * 100)
      return `${cardinalText(whole)} and ${String(cents).padStart(2, '0')}/100`
    }
    default:
      return null
  }
}

function fillPicture(section: string, value: number): string {
  const digits = /[0#][0#,]*(?:\.[0#]+)?/.exec(section)
  if (!digits) return section.replace(/'([^']*)'/g, '$1')
  const [whole, frac = ''] = digits[0].split('.')
  const fixed = Math.abs(value).toFixed(frac.length)
  const [intPart, fracPart = ''] = fixed.split('.')
  const minInt = (whole.match(/0/g) ?? []).length
  let int = intPart === '0' && minInt === 0 ? '' : intPart.padStart(minInt, '0')
  if (whole.includes(',')) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  let fracOut = ''
  for (let i = 0; i < frac.length; i++) fracOut += fracPart[i] ?? '0'
  // trailing # places drop their zeros
  for (let i = frac.length - 1; i >= 0 && frac[i] === '#' && fracOut.endsWith('0'); i--)
    fracOut = fracOut.slice(0, -1)
  const number = fracOut ? `${int || '0'}.${fracOut}` : int || (minInt === 0 ? '0' : int)
  const literal = (text: string) => text.replace(/'([^']*)'/g, '$1')
  return (
    literal(section.slice(0, digits.index)) +
    number +
    literal(section.slice(digits.index + digits[0].length))
  )
}

/** a `\#` picture: `positive;negative;zero` sections of 0 / # places, a comma groups thousands */
export function formatNumericPicture(value: number, picture: string): string {
  const sections = picture.split(';')
  if (value < 0) {
    return sections[1] !== undefined
      ? fillPicture(sections[1], value)
      : `-${fillPicture(sections[0], value)}`
  }
  if (value === 0 && sections[2] !== undefined) return fillPicture(sections[2], value)
  return fillPicture(sections[0], value)
}

/** a numeric result the way the code's `\#` and `\*` switches ask for it */
export function formatFieldNumber(code: string, value: number): string {
  const picture = numericPictureOf(code)
  const format = formatSwitchOf(code)
  if (picture) return formatFieldText(code, formatNumericPicture(value, picture))
  const shaped = format ? formatNumberAs(value, format) : null
  return shaped ?? formatFieldText(code, String(value))
}

/** a text result with the code's capitalization switch applied */
export function formatFieldText(code: string, value: string): string {
  switch (formatSwitchOf(code)?.toLowerCase()) {
    case 'upper':
      return value.toUpperCase()
    case 'lower':
      return value.toLowerCase()
    case 'firstcap':
      return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase()
    case 'caps':
      return value
        .toLowerCase()
        .replace(/(^|[\s-])(\p{L})/gu, (_, sep: string, ch: string) => sep + ch.toUpperCase())
    default:
      return value
  }
}
