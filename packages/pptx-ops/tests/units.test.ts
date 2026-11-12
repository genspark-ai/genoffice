/**
 * Unit-suffix parsing for length fields ("2.54cm", "1in", "12pt", "96px").
 * Lengths are document-space EMU unless the field name says otherwise; a bare
 * number stays EMU. The executor normalizes suffix strings in place before
 * validation, so every op keeps seeing integers.
 */
import { describe, it, expect } from 'vitest'
import { normalizeLengthUnits, parseLength } from '../src/ops/units'

describe('parseLength conversions', () => {
  it('maps one inch expressed in every unit to 914400 EMU', () => {
    expect(parseLength('1in')).toBe(914400)
    expect(parseLength('2.54cm')).toBe(914400)
    expect(parseLength('25.4mm')).toBe(914400)
    expect(parseLength('72pt')).toBe(914400)
    expect(parseLength('6pc')).toBe(914400)
    expect(parseLength('96px')).toBe(914400)
    expect(parseLength('914400emu')).toBe(914400)
  })

  it('converts the task prompt examples exactly', () => {
    expect(parseLength('2.54cm')).toBe(914400)
    expect(parseLength('1in')).toBe(914400)
    expect(parseLength('12pt')).toBe(12 * 12700)
    expect(parseLength('96px')).toBe(96 * 9525)
  })

  it('handles decimals, negatives, case and surrounding whitespace', () => {
    expect(parseLength('0.5in')).toBe(457200)
    expect(parseLength('-1in')).toBe(-914400)
    expect(parseLength('1IN')).toBe(914400)
    expect(parseLength('2.54CM')).toBe(914400)
    expect(parseLength('  1in  ')).toBe(914400)
    expect(parseLength('10mm')).toBe(360000)
    expect(parseLength('1pc')).toBe(152400)
  })

  it('rounds fractional EMU to an integer', () => {
    expect(parseLength('0.1mm')).toBe(Math.round(0.1 * 36000))
  })

  it('returns undefined for bare numbers and unknown units', () => {
    expect(parseLength('123')).toBeUndefined()
    expect(parseLength('')).toBeUndefined()
    expect(parseLength('abc')).toBeUndefined()
    expect(parseLength('1kg')).toBeUndefined()
    expect(parseLength('1 in ch')).toBeUndefined()
  })

  it('returns undefined when the EMU conversion overflows to Infinity', () => {
    // 400 digits overflow Number(), so the EMU product is Infinity; Infinity
    // must not reach the XML writer or a numeric field.
    expect(parseLength(`${'9'.repeat(400)}cm`)).toBeUndefined()
    expect(parseLength(`${'9'.repeat(400)}in`)).toBeUndefined()
  })
})

describe('normalizeLengthUnits', () => {
  it('normalizes geometry keys and every ...Emu key', () => {
    expect(normalizeLengthUnits('1in', 'x')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'y')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'cx')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'cy')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'dx')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'dy')).toBe(914400)
    expect(normalizeLengthUnits('12pt', 'widthEmu')).toBe(152400)
    expect(normalizeLengthUnits('2.54cm', 'colWidthsEmu')).toBe(914400)
  })

  it('leaves non-length keys and keyless strings untouched', () => {
    expect(normalizeLengthUnits('1in', 'color')).toBe('1in')
    expect(normalizeLengthUnits('1in')).toBe('1in')
    expect(normalizeLengthUnits('hello', 'x')).toBe('hello')
    expect(normalizeLengthUnits(914400, 'x')).toBe(914400)
  })

  it('leaves an overflowing length as typed so the schema rejects it', () => {
    const huge = `${'9'.repeat(400)}cm`
    expect(normalizeLengthUnits(huge, 'x')).toBe(huge)
  })

  it('only normalizes inset sides under insets/insetsEmu', () => {
    expect(normalizeLengthUnits('1in', 'l', 'insets')).toBe(914400)
    expect(normalizeLengthUnits('1in', 't', 'insetsEmu')).toBe(914400)
    expect(normalizeLengthUnits('1in', 'l', 'other')).toBe('1in')
    expect(normalizeLengthUnits('1in', 'l')).toBe('1in')
  })

  it('walks arrays and nested objects', () => {
    expect(normalizeLengthUnits(['1in', '2in'], 'x')).toEqual([914400, 1828800])
    expect(normalizeLengthUnits({ x: '1in', color: '1in' })).toEqual({
      x: 914400,
      color: '1in',
    })
    expect(normalizeLengthUnits({ insets: { l: '1in', t: '12pt' } })).toEqual({
      insets: { l: 914400, t: 152400 },
    })
  })

  it('passes Uint8Array payloads through by reference', () => {
    const bytes = new Uint8Array([1, 2, 3])
    expect(normalizeLengthUnits(bytes, 'bytes')).toBe(bytes)
  })
})
