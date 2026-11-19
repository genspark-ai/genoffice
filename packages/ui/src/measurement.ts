/**
 * Physical length units shared by every editor: Word's Preferences ▸ General ▸
 * "Show measurements in units of", applied app-wide. One preference, held by
 * the shell (Settings ▸ General) and handed to each editor; the conversions
 * here are pure so the main process and every renderer agree on them.
 *
 * Lengths go through twips (1/1440 in). Editors that store EMU (slides, sheet
 * drawings) or CSS px use the adapters at the bottom. Font sizes and line
 * spacing stay in points whatever the setting, as in Office.
 */
export type MeasurementUnit = 'in' | 'cm' | 'mm' | 'pt' | 'pi'

export const MEASUREMENT_UNITS: readonly MeasurementUnit[] = ['in', 'cm', 'mm', 'pt', 'pi']

export const TWIPS_PER: Record<MeasurementUnit, number> = {
  in: 1440,
  cm: 1440 / 2.54,
  mm: 144 / 2.54,
  pt: 20,
  pi: 240,
}

export const DECIMALS: Record<MeasurementUnit, number> = { in: 2, cm: 2, mm: 1, pt: 1, pi: 2 }

/** Arrow-key / spinner increment per unit. */
export const UNIT_STEP: Record<MeasurementUnit, number> = {
  in: 0.1,
  cm: 0.1,
  mm: 1,
  pt: 1,
  pi: 0.5,
}

export const UNIT_SUFFIX: Record<MeasurementUnit, string> = {
  in: '"',
  cm: 'cm',
  mm: 'mm',
  pt: 'pt',
  pi: 'pi',
}

export const isMeasurementUnit = (v: unknown): v is MeasurementUnit =>
  typeof v === 'string' && (MEASUREMENT_UNITS as readonly string[]).includes(v)

const INCH_REGIONS = new Set(['US', 'LR', 'MM'])

/** Metric everywhere except the three inch countries; a bare `en` reads as en-US. */
export function defaultMeasurementUnit(locale: string): MeasurementUnit {
  let region: string | undefined
  try {
    region = new Intl.Locale(locale).maximize().region
  } catch {
    region = /-([A-Za-z]{2})(?:-|$)/.exec(locale)?.[1]
  }
  if (region) return INCH_REGIONS.has(region.toUpperCase()) ? 'in' : 'cm'
  return /^en\b/i.test(locale) ? 'in' : 'cm'
}

/**
 * Default for an ISO 3166 region code, e.g. the operating system's region
 * (Electron's app.getLocaleCountryCode()). That beats the UI language: an
 * English UI in Indonesia is still a metric user. null when the code is unusable.
 */
export function measurementUnitForRegion(
  region: string | null | undefined,
): MeasurementUnit | null {
  if (!region || !/^[A-Za-z]{2}$/.test(region)) return null
  return INCH_REGIONS.has(region.toUpperCase()) ? 'in' : 'cm'
}

export const twipsPerUnit = (unit: MeasurementUnit): number => TWIPS_PER[unit]

export function toUnit(twips: number, unit: MeasurementUnit): number {
  const f = 10 ** DECIMALS[unit]
  return Math.round((twips / TWIPS_PER[unit]) * f) / f
}

export function fromUnit(value: number, unit: MeasurementUnit): number {
  return Math.round(value * TWIPS_PER[unit])
}

/** `2.54 cm`, `1"`, `12 pt`: rounded per unit, trailing zeros trimmed, the
 * inch mark glued to the number like Word. */
export function formatLength(twips: number, unit: MeasurementUnit, suffix = UNIT_SUFFIX[unit]) {
  const n = toUnit(twips, unit)
  const text = n.toFixed(DECIMALS[unit]).replace(/\.?0+$/, '')
  const num = text === '' || text === '-' || text === '-0' ? '0' : text
  return suffix === '"' ? `${num}"` : `${num} ${suffix}`
}

const SUFFIX_UNITS: Record<string, MeasurementUnit | 'px'> = {
  '"': 'in',
  '″': 'in',
  in: 'in',
  inch: 'in',
  inches: 'in',
  cm: 'cm',
  mm: 'mm',
  pt: 'pt',
  pts: 'pt',
  point: 'pt',
  points: 'pt',
  pi: 'pi',
  pica: 'pi',
  picas: 'pi',
  px: 'px',
}

/** Parses `2cm`, `1"`, `1 in`, `36pt`, `3pi`, `2,5 cm`, a bare number (in
 * `unit`) or a localized suffix from `localSuffixes`; null when unreadable. */
export function parseLength(
  text: string,
  unit: MeasurementUnit,
  localSuffixes?: Partial<Record<MeasurementUnit, string>>,
): number | null {
  const m = /^\s*([+-]?(?:\d+[.,]?\d*|[.,]\d+))\s*(.*?)\s*$/.exec(text)
  if (!m) return null
  const value = Number((m[1] ?? '').replace(',', '.'))
  if (!Number.isFinite(value)) return null
  const suffix = (m[2] ?? '').toLowerCase()
  if (suffix === '') return fromUnit(value, unit)
  if (localSuffixes) {
    for (const u of MEASUREMENT_UNITS) {
      const local = localSuffixes[u]?.toLowerCase()
      if (local && local === suffix) return fromUnit(value, u)
    }
  }
  const target = SUFFIX_UNITS[suffix]
  if (!target) return null
  return target === 'px' ? Math.round(value * 15) : fromUnit(value, target)
}

/** CSS px at 96 dpi: 1 px = 15 twips */
export const pxToTwips = (px: number): number => Math.round(px * 15)
export const twipsToPx = (twips: number): number => twips / 15

/** EMU (DrawingML): 1 twip = 635 EMU */
export const emuToTwips = (emu: number): number => emu / 635
export const twipsToEmu = (twips: number): number => Math.round(twips * 635)

/**
 * The app-wide setting as the shell hands it out: the unit in effect, and
 * whether the user picked it (false = following the system region).
 */
export interface MeasurementUnitState {
  unit: MeasurementUnit
  explicit: boolean
}

export const isMeasurementUnitState = (v: unknown): v is MeasurementUnitState =>
  !!v &&
  typeof v === 'object' &&
  isMeasurementUnit((v as MeasurementUnitState).unit) &&
  typeof (v as MeasurementUnitState).explicit === 'boolean'

/** IPC channels every editor's preload uses (shell main answers them) */
export const MEASUREMENT_UNIT_CHANNELS = {
  get: 'app:get-measurement-unit',
  set: 'app:set-measurement-unit',
  changed: 'app:measurement-unit-changed',
} as const
