/** Word's date-time picture (the text after a DATE or TIME field's \@ switch), English names. */

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
]
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

const two = (n: number): string => String(n).padStart(2, '0')

// longest first, so MMMM is not read as MM MM; a quoted run is literal text
const TOKEN = /'[^']*'|am\/pm|AM\/PM|M{1,4}|d{1,4}|yyyy|yy|h{1,2}|H{1,2}|m{1,2}|s{1,2}/g

const isEnglish = (lang?: string): boolean => !lang || lang.toLowerCase().startsWith('en')

/** one named part of the date in `lang`, the way that language writes it standing in a date */
function localName(
  date: Date,
  lang: string,
  type: 'month' | 'weekday' | 'dayPeriod',
  width: 'long' | 'short',
): string | null {
  try {
    const options: Intl.DateTimeFormatOptions =
      type === 'month'
        ? { month: width, day: 'numeric' }
        : type === 'weekday'
          ? { weekday: width }
          : { hour: 'numeric', hour12: true }
    const parts = new Intl.DateTimeFormat(lang, { ...options, calendar: 'gregory' }).formatToParts(
      date,
    )
    return parts.find((p) => p.type === type)?.value ?? null
  } catch {
    return null
  }
}

/** ASCII as its full-width forms, as Word's \* DBCHAR switch shows a result */
export function toFullWidth(text: string): string {
  return text.replace(/[!-~]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) + 0xfee0))
}

/** whether a field instruction asks for a full-width result */
export const wantsFullWidth = (instr: string): boolean => /\\\*\s*DBCHAR\b/i.test(instr)

/** `lang` names the months, days and day halves; English when left out */
export function formatDatePicture(date: Date, picture: string, lang?: string): string {
  const hour12 = date.getHours() % 12 || 12
  const named = (type: 'month' | 'weekday' | 'dayPeriod', width: 'long' | 'short') =>
    isEnglish(lang) ? null : localName(date, lang!, type, width)
  return picture.replace(TOKEN, (token) => {
    switch (token) {
      case 'M':
        return String(date.getMonth() + 1)
      case 'MM':
        return two(date.getMonth() + 1)
      case 'MMM':
        return named('month', 'short') ?? MONTHS[date.getMonth()].slice(0, 3)
      case 'MMMM':
        return named('month', 'long') ?? MONTHS[date.getMonth()]
      case 'd':
        return String(date.getDate())
      case 'dd':
        return two(date.getDate())
      case 'ddd':
        return named('weekday', 'short') ?? DAYS[date.getDay()].slice(0, 3)
      case 'dddd':
        return named('weekday', 'long') ?? DAYS[date.getDay()]
      case 'yy':
        return two(date.getFullYear() % 100)
      case 'yyyy':
        return String(date.getFullYear())
      case 'h':
        return String(hour12)
      case 'hh':
        return two(hour12)
      case 'H':
        return String(date.getHours())
      case 'HH':
        return two(date.getHours())
      case 'm':
        return String(date.getMinutes())
      case 'mm':
        return two(date.getMinutes())
      case 's':
        return String(date.getSeconds())
      case 'ss':
        return two(date.getSeconds())
      case 'am/pm':
      case 'AM/PM':
        return named('dayPeriod', 'long') ?? (date.getHours() < 12 ? 'AM' : 'PM')
      default:
        return token.slice(1, -1)
    }
  })
}

/** the picture of a field instruction such as `DATE \@ "MMMM d, yyyy"`, or null without one */
export function datePictureOf(instr: string): string | null {
  const m = /\\@\s*(?:"([^"]*)"|(\S+))/.exec(instr)
  return m ? (m[1] ?? m[2]) : null
}

/** the formats Word's Date and Time dialog lists for English (US), in its order */
export const DATE_TIME_PICTURES = [
  'M/d/yy',
  'dddd, MMMM d, yyyy',
  'MMMM d, yyyy',
  'M/d/yyyy',
  'yyyy-MM-dd',
  'd-MMM-yy',
  'M.d.yy',
  'MMM. d, yy',
  'd MMMM yyyy',
  'MMMM yy',
  'MMM-yy',
  'M/d/yy h:mm am/pm',
  'M/d/yy h:mm:ss am/pm',
  'h:mm am/pm',
  'h:mm:ss am/pm',
  'HH:mm',
  'HH:mm:ss',
] as const

const PROBE = new Date(2001, 1, 3, 16, 5, 7)

/** the picture that writes a date the way `lang` does in one of Intl's styles; null when it cannot be told */
function pictureOfStyle(lang: string, options: Intl.DateTimeFormatOptions): string | null {
  try {
    const format = new Intl.DateTimeFormat(lang, {
      ...options,
      numberingSystem: 'latn',
      calendar: 'gregory',
    })
    const longMonth = localName(PROBE, lang, 'month', 'long')
    return format
      .formatToParts(PROBE)
      .map(({ type, value }) => {
        const wide = value.length > 1 && value.startsWith('0')
        switch (type) {
          case 'year':
            return value.length > 2 ? 'yyyy' : 'yy'
          case 'month':
            return /^\d+$/.test(value) ? (wide ? 'MM' : 'M') : value === longMonth ? 'MMMM' : 'MMM'
          case 'day':
            return wide ? 'dd' : 'd'
          case 'weekday':
            return value === localName(PROBE, lang, 'weekday', 'long') ? 'dddd' : 'ddd'
          case 'hour':
            return value === '16' ? 'HH' : wide ? 'hh' : 'h'
          case 'minute':
            return 'mm'
          case 'second':
            return 'ss'
          case 'dayPeriod':
            return 'am/pm'
          case 'literal':
            // text a picture would read as a token is quoted
            return /[A-Za-z]/.test(value) ? `'${value.replace(/'/g, '')}'` : value
          default:
            throw new Error(type)
        }
      })
      .join('')
  } catch {
    return null
  }
}

const STYLES: Intl.DateTimeFormatOptions[] = [
  { dateStyle: 'short' },
  { dateStyle: 'full' },
  { dateStyle: 'long' },
  { dateStyle: 'medium' },
  { dateStyle: 'short', timeStyle: 'short' },
  { dateStyle: 'short', timeStyle: 'medium' },
  { timeStyle: 'short' },
  { timeStyle: 'medium' },
]

/**
 * The formats the Date and Time dialog lists for a language: Word's own list for
 * English, and for another language the ways that language writes a date.
 */
export function dateTimePicturesFor(lang: string): readonly string[] {
  if (lang.toLowerCase() === 'en-us') return DATE_TIME_PICTURES
  const out = new Set<string>()
  for (const [i, style] of STYLES.entries()) {
    const picture = pictureOfStyle(lang, style)
    if (picture) out.add(picture)
    if (i === 3) out.add('yyyy-MM-dd')
  }
  out.add('HH:mm')
  out.add('HH:mm:ss')
  return out.size > 3 ? [...out] : DATE_TIME_PICTURES
}
