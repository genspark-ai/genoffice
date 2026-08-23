/// Excel renders builtin numFmtIds 14/22 — the asterisked short-date entries
/// in the Format Cells dialog — per the OS regional short-date setting, not a
/// fixed pattern. The region drives the pattern, not the UI language: an
/// English UI on a yyyy/m/d-region machine still shows yyyy/m/d.

export const DEFAULT_SHORT_DATE = 'm/d/yyyy'

const PART_TOKENS: Record<string, ((value: string) => string) | undefined> = {
  year: (value) => (value.length === 2 ? 'yy' : 'yyyy'),
  month: (value) => (value.startsWith('0') ? 'mm' : 'm'),
  day: (value) => (value.startsWith('0') ? 'dd' : 'd'),
}

export function shortDatePatternForSystemLocale(systemLocale: string): string {
  try {
    const region = new Intl.Locale(systemLocale).region ?? 'US'
    // Language must match the region (CN→zh, DE→de): keeping the system tag's
    // language ("en-CN") or a script subtag makes Intl fall back to bare "en".
    const language = new Intl.Locale('und', { region }).maximize().language
    const parts = new Intl.DateTimeFormat(`${language}-${region}`, {
      calendar: 'gregory',
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
    }).formatToParts(new Date(2016, 2, 9))
    const pattern = parts
      .map((part) => PART_TOKENS[part.type]?.(part.value) ?? part.value.replace(/[^./\- ,]/g, ''))
      .join('')
    return /^(?=.*y)(?=.*m)(?=.*d)[ymd./\- ,]+$/.test(pattern) ? pattern : DEFAULT_SHORT_DATE
  } catch {
    return DEFAULT_SHORT_DATE
  }
}

let systemShortDate = DEFAULT_SHORT_DATE

export function setSystemShortDate(pattern: string): void {
  systemShortDate = pattern
}

export function getSystemShortDate(): string {
  return systemShortDate
}

/// Save-side inverse: a pattern equal to the system short date came from
/// builtin 14 (or 22 with time), so writing the id back keeps the cell
/// locale-reactive in Excel.
export function shortDateNumFmtId(pattern: string): number | undefined {
  const shortDate = getSystemShortDate()
  if (pattern === shortDate) return 14
  if (pattern === `${shortDate} h:mm`) return 22
  return undefined
}
