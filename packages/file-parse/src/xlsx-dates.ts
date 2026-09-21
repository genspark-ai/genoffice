/** Which calendar parts a number format renders; null when it is not a date/time format. */
export interface DateFormatParts {
  date: boolean
  time: boolean
  seconds: boolean
  /** [h] / [mm] / [ss] elapsed-duration tokens, not clock time */
  elapsed: boolean
}

const DATE_ONLY: DateFormatParts = { date: true, time: false, seconds: false, elapsed: false }
const DATE_TIME: DateFormatParts = { date: true, time: true, seconds: false, elapsed: false }
const TIME_ONLY: DateFormatParts = { date: false, time: true, seconds: false, elapsed: false }
const TIME_SECONDS: DateFormatParts = { date: false, time: true, seconds: true, elapsed: false }
const ELAPSED: DateFormatParts = { date: false, time: true, seconds: true, elapsed: true }

/**
 * Built-in ids (ECMA-376 §18.8.30 plus the CJK 27-36 / 50-58 ranges Excel reserves) that
 * render as dates or times. Everything else built-in is numeric or text.
 */
const BUILTIN: ReadonlyMap<number, DateFormatParts> = new Map<number, DateFormatParts>([
  [14, DATE_ONLY],
  [15, DATE_ONLY],
  [16, DATE_ONLY],
  [17, DATE_ONLY],
  [18, TIME_ONLY],
  [19, TIME_SECONDS],
  [20, TIME_ONLY],
  [21, TIME_SECONDS],
  [22, DATE_TIME],
  ...[27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58].map(
    (id): [number, DateFormatParts] => [id, DATE_ONLY],
  ),
  [32, TIME_ONLY],
  [33, TIME_SECONDS],
  [45, TIME_SECONDS],
  [46, ELAPSED],
  [47, TIME_SECONDS],
])

export function builtinDateFormat(numFmtId: number): DateFormatParts | null {
  return BUILTIN.get(numFmtId) ?? null
}

/**
 * Classify a custom formatCode. Only the first section (positive numbers) decides; quoted
 * literals, escaped characters, fill/skip tokens and colour/locale conditions are not tokens.
 */
export function classifyFormatCode(code: string): DateFormatParts | null {
  const section = code.split(';')[0] ?? ''
  if (/general/i.test(section) && !/[ydhs]/i.test(section.replace(/general/gi, ''))) return null
  let elapsed = false
  let elapsedMinutes = false
  const stripped = section
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    // elapsed [h] / [mm] / [ss] keep their letter so the minute adjacency rule below still sees them
    .replace(/\[(h+|m+|s+)\]/gi, (_all, token: string) => {
      elapsed = true
      if (/^m/i.test(token)) elapsedMinutes = true
      return token
    })
    .replace(/\[[^\]]*\]/g, '')
    .toLowerCase()
  const hasAmPm = /am\/pm|a\/p/.test(stripped)
  const body = stripped.replace(/am\/pm|a\/p/g, '')
  const hasY = body.includes('y')
  const hasD = body.includes('d')
  const hasH = body.includes('h') || hasAmPm
  const hasS = body.includes('s')
  const hasM = body.includes('m')
  if (!hasY && !hasD && !hasH && !hasS && !hasM) return null
  // 'm' is a month unless it sits next to hours or seconds (Excel's own rule): mmm alone is a month
  const minuteM = hasM && (elapsedMinutes || /h\s*[:.]?\s*m|m\s*[:.]?\s*s/.test(body))
  const date = hasY || hasD || (hasM && !minuteM)
  const time = hasH || hasS || minuteM || elapsed
  if (!date && !time) return null
  return { date, time, seconds: hasS, elapsed: elapsed && !date }
}

function pad(n: number, width = 2): string {
  return String(n).padStart(width, '0')
}

/**
 * Render an Excel serial in ISO-like text. 1900 system: serial 60 is the phantom
 * 1900-02-29 (Lotus compatibility) and 1–59 sit one day before the linear rule.
 * Returns null when the value cannot be a calendar value under this format.
 */
export function formatSerial(
  serial: number,
  parts: DateFormatParts,
  date1904: boolean,
): string | null {
  if (!Number.isFinite(serial)) return null
  if (parts.elapsed) {
    if (serial < 0) return null
    const total = Math.round(serial * 86400)
    const h = Math.floor(total / 3600)
    const m = Math.floor((total % 3600) / 60)
    const s = total % 60
    return `${h}:${pad(m)}:${pad(s)}`
  }
  if (serial < 0) return null
  let days = Math.floor(serial)
  let secs = Math.round((serial - days) * 86400)
  if (secs >= 86400) {
    days += 1
    secs = 0
  }
  const timeText = parts.time
    ? `${pad(Math.floor(secs / 3600))}:${pad(Math.floor((secs % 3600) / 60))}${parts.seconds ? `:${pad(secs % 60)}` : ''}`
    : ''
  if (!parts.date) return timeText
  let dateText: string
  if (!date1904 && days === 60) {
    dateText = '1900-02-29'
  } else {
    if (!date1904 && days === 0) return null
    // 1900 system: 1899-12-30 as epoch makes serials ≥ 61 land right; below 60 shift back one day
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
    const shift = !date1904 && days < 60 ? 1 : 0
    const d = new Date(epoch + (days + shift) * 86400000)
    dateText = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`
  }
  return timeText ? `${dateText} ${timeText}` : dateText
}
