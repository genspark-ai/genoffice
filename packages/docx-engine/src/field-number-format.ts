import { formatNumber } from './list-markers'

/**
 * The `\*` number formats of a field (ECMA-376 Part 1, 17.16.4.3) that the
 * standard pairs with an ST_NumberFormat value, by the switch's lower case.
 * `max` is the last value the standard's own examples settle; `plain` keeps to
 * numbers written without a zero between digits, which the examples do not show.
 */
const COUNTED: Readonly<Record<string, { numFmt: string; max: number; plain?: boolean }>> = {
  chinesenum1: { numFmt: 'chineseCounting', max: 99 },
  chinesenum2: { numFmt: 'chineseLegalSimplified', max: 9999, plain: true },
  chinesenum3: { numFmt: 'chineseCountingThousand', max: 9999, plain: true },
  // "sequential digital ideographs": the standard names no value, its example is DBNUM1's
  kanjinum1: { numFmt: 'ideographDigital', max: 999_999 },
  kanjinum2: { numFmt: 'japaneseCounting', max: 99 },
  dbnum1: { numFmt: 'ideographDigital', max: 999_999 },
  dbnum4: { numFmt: 'japaneseDigitalTenThousand', max: 999_999 },
  aiueo: { numFmt: 'aiueo', max: 46 },
  zodiac1: { numFmt: 'ideographTraditional', max: 10 },
  zodiac2: { numFmt: 'ideographZodiac', max: 12 },
  ganada: { numFmt: 'ganada', max: 14 },
  chosung: { numFmt: 'chosung', max: 14 },
}

const WORDED = [
  'arabic',
  'arabicdash',
  'alphabetic',
  'roman',
  'ordinal',
  'cardtext',
  'ordtext',
  'hex',
  'dollartext',
  'zodiac3',
]

/** a `\*` argument that shapes a number and that `fieldNumberText` has a rule for */
export const isFieldNumberFormat = (name: string): boolean => {
  const key = name.toLowerCase()
  return key in COUNTED || WORDED.includes(key)
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

/**
 * A number as a field's `\*` format shows it; null where this has no rule
 * (a format it does not know, a value outside what it computes), so the caller
 * leaves the result to Word. The worded formats are English.
 */
export function fieldNumberText(n: number, format: string): string | null {
  const key = format.toLowerCase()
  const whole = Math.trunc(Math.abs(n))
  const inWords = whole < 1_000_000
  // the standard's two spellings: `Roman` / `ALPHABETIC` upper case, `roman` / `alphabetic` lower
  const upper = /^[A-Z]/.test(format)
  const counted = COUNTED[key]
  if (counted) {
    if (!Number.isInteger(n) || n < 1 || n > counted.max) return null
    if (counted.plain && /0[1-9]/.test(String(n))) return null
    return formatNumber(n, counted.numFmt)
  }
  switch (key) {
    case 'arabic':
      return String(n)
    case 'arabicdash':
      return `- ${n} -`
    case 'alphabetic': {
      if (whole < 1) return null
      const letters = formatNumber(whole, 'upperLetter')
      return upper ? letters : letters.toLowerCase()
    }
    case 'roman': {
      if (whole < 1 || whole >= 4000) return null
      const numerals = formatNumber(whole, 'upperRoman')
      return upper ? numerals : numerals.toLowerCase()
    }
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
    case 'zodiac3':
      // the sixty-pair cycle: the stems and the branches each going round
      if (!Number.isInteger(n) || n < 1 || n > 60) return null
      return (
        formatNumber(((n - 1) % 10) + 1, 'ideographTraditional') +
        formatNumber(((n - 1) % 12) + 1, 'ideographZodiac')
      )
    default:
      return null
  }
}
