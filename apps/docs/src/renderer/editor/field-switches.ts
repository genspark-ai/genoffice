import { fieldNumberText } from '@genoffice/docx-engine'

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
  const shaped = format ? fieldNumberText(value, format) : null
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
