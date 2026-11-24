import { describe, expect, it } from 'vitest'

import {
  defaultMeasurementUnit,
  emuToTwips,
  formatLength,
  isMeasurementUnitState,
  parseLength,
  parseLengthEmu,
  twipsToEmu,
} from '../src/measurement'

describe('measurement units', () => {
  it('defaults by the UI language: inches for English, metric elsewhere', () => {
    expect(defaultMeasurementUnit('en')).toBe('in')
    expect(defaultMeasurementUnit('zh')).toBe('cm')
    expect(defaultMeasurementUnit('pt')).toBe('cm')
  })

  it('reads the region of a full locale', () => {
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

  it('parses straight to exact EMU for DrawingML lengths', () => {
    expect(parseLengthEmu('10', 'mm')).toBe(360000)
    expect(parseLengthEmu('0,5 cm', 'pt')).toBe(180000)
    expect(parseLengthEmu('3.5', 'mm')).toBe(126000)
    expect(parseLengthEmu('1"', 'cm')).toBe(914400)
    expect(parseLengthEmu('72 pt', 'cm')).toBe(914400)
    expect(parseLengthEmu('96px', 'cm')).toBe(914400)
    expect(parseLengthEmu('furlong', 'cm')).toBeNull()
  })

  it('validates the state the shell hands out', () => {
    expect(isMeasurementUnitState({ unit: 'cm', explicit: false })).toBe(true)
    expect(isMeasurementUnitState({ unit: 'px', explicit: true })).toBe(false)
    expect(isMeasurementUnitState({ unit: 'cm' })).toBe(false)
    expect(isMeasurementUnitState(null)).toBe(false)
  })
})
