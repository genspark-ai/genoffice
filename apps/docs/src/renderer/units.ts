/** Word's Preferences ▸ General ▸ "Show measurements in units of": one
 * preference every length field (ruler, margins, indents, sizes) follows.
 * Storage stays in twips; only presentation changes. Font size, line and
 * paragraph spacing stay in points, as in Word. */
export type MeasurementUnit = 'in' | 'cm' | 'mm' | 'pt' | 'pi'

export const MEASUREMENT_UNITS: readonly MeasurementUnit[] = ['in', 'cm', 'mm', 'pt', 'pi']
export const MEASUREMENT_UNIT_KEY = 'aidocs.measurementUnit'
const CHANGE_EVENT = 'aidocs:measurement-unit'

const TWIPS_PER: Record<MeasurementUnit, number> = {
  in: 1440,
  cm: 1440 / 2.54,
  mm: 144 / 2.54,
  pt: 20,
  pi: 240,
}

const DECIMALS: Record<MeasurementUnit, number> = { in: 2, cm: 2, mm: 1, pt: 1, pi: 2 }

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

export function storedMeasurementUnit(): MeasurementUnit | null {
  try {
    const raw = globalThis.localStorage?.getItem(MEASUREMENT_UNIT_KEY)
    return isMeasurementUnit(raw) ? raw : null
  } catch {
    return null
  }
}

export function measurementUnit(): MeasurementUnit {
  return (
    storedMeasurementUnit() ?? defaultMeasurementUnit(globalThis.navigator?.language ?? 'en-US')
  )
}

export function setMeasurementUnit(unit: MeasurementUnit | null): void {
  try {
    if (unit) globalThis.localStorage?.setItem(MEASUREMENT_UNIT_KEY, unit)
    else globalThis.localStorage?.removeItem(MEASUREMENT_UNIT_KEY)
  } catch {
    /* storage unavailable: the choice just doesn't persist */
  }
  globalThis.dispatchEvent?.(new Event(CHANGE_EVENT))
}

export function subscribeMeasurementUnit(listener: () => void): () => void {
  window.addEventListener(CHANGE_EVENT, listener)
  window.addEventListener('storage', listener)
  return () => {
    window.removeEventListener(CHANGE_EVENT, listener)
    window.removeEventListener('storage', listener)
  }
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
  const value = Number(m[1].replace(',', '.'))
  if (!Number.isFinite(value)) return null
  const suffix = m[2].toLowerCase()
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

export const pxToTwips = (px: number): number => Math.round(px * 15)
export const twipsToPx = (twips: number): number => twips / 15
