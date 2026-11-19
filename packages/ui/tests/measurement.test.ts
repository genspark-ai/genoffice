import { describe, expect, it } from 'vitest'

import {
  defaultMeasurementUnit,
  emuToTwips,
  formatLength,
  isMeasurementUnitState,
  measurementUnitForRegion,
  parseLength,
  twipsToEmu,
} from '../src/measurement'

describe('measurement units', () => {
  it('defaults by the OS region, metric everywhere but the inch countries', () => {
    expect(measurementUnitForRegion('ID')).toBe('cm')
    expect(measurementUnitForRegion('gb')).toBe('cm')
    expect(measurementUnitForRegion('US')).toBe('in')
    expect(measurementUnitForRegion('MM')).toBe('in')
    expect(measurementUnitForRegion('')).toBeNull()
    expect(measurementUnitForRegion(undefined)).toBeNull()
    expect(measurementUnitForRegion('419')).toBeNull()
  })

  it('still reads a locale when no region is known', () => {
    expect(defaultMeasurementUnit('en-US')).toBe('in')
    expect(defaultMeasurementUnit('en-GB')).toBe('cm')
    expect(defaultMeasurementUnit('id-ID')).toBe('cm')
  })

  it('converts EMU through twips', () => {
    expect(emuToTwips(914400)).toBe(1440)
    expect(twipsToEmu(1440)).toBe(914400)
    expect(formatLength(emuToTwips(360000), 'cm')).toBe('1 cm')
  })

  it('formats and parses every unit', () => {
    expect(formatLength(1440, 'in')).toBe('1"')
    expect(formatLength(1440, 'mm')).toBe('25.4 mm')
    expect(formatLength(1440, 'pt')).toBe('72 pt')
    expect(formatLength(1440, 'pi')).toBe('6 pi')
    expect(parseLength('2,54 cm', 'in')).toBe(1440)
    expect(parseLength('6', 'pi')).toBe(1440)
  })

  it('validates the state the shell hands out', () => {
    expect(isMeasurementUnitState({ unit: 'cm', explicit: false })).toBe(true)
    expect(isMeasurementUnitState({ unit: 'px', explicit: true })).toBe(false)
    expect(isMeasurementUnitState({ unit: 'cm' })).toBe(false)
    expect(isMeasurementUnitState(null)).toBe(false)
  })
})
