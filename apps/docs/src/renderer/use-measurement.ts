import { useSyncExternalStore } from 'react'
import { useI18n, type StringKey } from './i18n/locale'
import {
  formatLength,
  fromUnit,
  measurementUnit,
  parseLength,
  subscribeMeasurementUnit,
  toUnit,
  type MeasurementUnit,
} from './units'

const ABBR_KEYS: Record<MeasurementUnit, StringKey> = {
  in: 'appUnitInAbbr',
  cm: 'appUnitCmAbbr',
  mm: 'appUnitMmAbbr',
  pt: 'appUnitPtAbbr',
  pi: 'appUnitPiAbbr',
}

export const UNIT_NAME_KEYS: Record<MeasurementUnit, StringKey> = {
  in: 'appUnitInches',
  cm: 'appUnitCentimeters',
  mm: 'appUnitMillimeters',
  pt: 'appUnitPoints',
  pi: 'appUnitPicas',
}

export interface Measurement {
  unit: MeasurementUnit
  /** localized short unit abbreviation */
  suffix: string
  format: (twips: number) => string
  parse: (text: string) => number | null
  toUnit: (twips: number) => number
  fromUnit: (value: number) => number
}

export function useMeasurementUnit(): MeasurementUnit {
  return useSyncExternalStore(subscribeMeasurementUnit, measurementUnit, measurementUnit)
}

/** Unit-bound helpers that follow the preference and the UI language. */
export function useMeasurement(): Measurement {
  const unit = useMeasurementUnit()
  const { t } = useI18n()
  const suffixes = {} as Record<MeasurementUnit, string>
  for (const u of Object.keys(ABBR_KEYS) as MeasurementUnit[]) suffixes[u] = t(ABBR_KEYS[u])
  return {
    unit,
    suffix: suffixes[unit],
    format: (twips: number) => formatLength(twips, unit, suffixes[unit]),
    parse: (text: string) => parseLength(text, unit, suffixes),
    toUnit: (twips: number) => toUnit(twips, unit),
    fromUnit: (value: number) => fromUnit(value, unit),
  }
}
