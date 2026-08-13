import { describe, expect, it } from 'vitest'
import {
  COLOR_PRESETS,
  hexTo255,
  hsvToRgb,
  isHexColor,
  rgb255ToHex,
  rgbToHsv,
} from '../src/renderer/color-presets'

describe('shared color presets', () => {
  it('every preset is a normalized #RRGGBB value', () => {
    for (const hex of COLOR_PRESETS) {
      expect(hex).toMatch(/^#[0-9A-F]{6}$/)
      expect(isHexColor(hex)).toBe(true)
    }
  })

  it('presets are unique', () => {
    expect(new Set(COLOR_PRESETS).size).toBe(COLOR_PRESETS.length)
  })
})

describe('isHexColor', () => {
  it('accepts 6-digit hex in either case', () => {
    expect(isHexColor('#a1B2c3')).toBe(true)
    expect(isHexColor('#000000')).toBe(true)
  })

  it('rejects everything else', () => {
    expect(isHexColor('a1b2c3')).toBe(false)
    expect(isHexColor('#fff')).toBe(false)
    expect(isHexColor('#12345g')).toBe(false)
    expect(isHexColor('#1234567')).toBe(false)
    expect(isHexColor('')).toBe(false)
  })
})

describe('hsv <-> rgb conversions', () => {
  it('maps the primary anchors', () => {
    expect(hsvToRgb(0, 0, 0)).toEqual([0, 0, 0])
    expect(hsvToRgb(0, 0, 1)).toEqual([255, 255, 255])
    expect(hsvToRgb(0, 1, 1)).toEqual([255, 0, 0])
    expect(hsvToRgb(120, 1, 1)).toEqual([0, 255, 0])
    expect(hsvToRgb(240, 1, 1)).toEqual([0, 0, 255])
    expect(rgbToHsv(255, 0, 0)).toEqual([0, 1, 1])
    expect(rgbToHsv(0, 0, 255)).toEqual([240, 1, 1])
  })

  it('round-trips every shared preset through hex -> hsv -> hex', () => {
    for (const hex of COLOR_PRESETS) {
      const [h, s, v] = rgbToHsv(...hexTo255(hex))
      expect(rgb255ToHex(hsvToRgb(h, s, v)).toUpperCase()).toBe(hex)
    }
  })

  it('hue stays in [0, 360)', () => {
    for (let r = 0; r <= 255; r += 51)
      for (let g = 0; g <= 255; g += 51)
        for (let b = 0; b <= 255; b += 51) {
          const [h, s, v] = rgbToHsv(r, g, b)
          expect(h).toBeGreaterThanOrEqual(0)
          expect(h).toBeLessThan(360)
          expect(s).toBeGreaterThanOrEqual(0)
          expect(s).toBeLessThanOrEqual(1)
          expect(v).toBeGreaterThanOrEqual(0)
          expect(v).toBeLessThanOrEqual(1)
        }
  })
})
