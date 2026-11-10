import { describe, expect, it } from 'vitest'
import { applyTint, normalizeStyleColor, resolveStyleColor } from '../src/domain/style-color'

const HEX_PATTERN = /^#[0-9A-F]{6}$/

function expectCleanHex(value: string): void {
  expect(value).toMatch(HEX_PATTERN)
  expect(value).not.toContain('NaN')
  expect(value).not.toContain('NAN')
  for (const channel of [value.slice(1, 3), value.slice(3, 5), value.slice(5, 7)]) {
    expect(Number.isNaN(parseInt(channel, 16))).toBe(false)
  }
}

describe('style-color hex guard', () => {
  it('keeps valid conversions unchanged', () => {
    expect(normalizeStyleColor('#ff0000')).toBe('#FF0000')
    expect(resolveStyleColor('#ff0000')).toBe('#FF0000')
    expect(applyTint('#FF0000', 0)).toBe('#FF0000')
    expect(applyTint('#000000', 0.5)).toBe('#7F7F7F')
    expect(applyTint('#FFFFFF', -0.5)).toBe('#7F7F7F')
    expect(applyTint('#4472C4', 0.5)).toBe('#A1B8E1')
    expect(applyTint('#4472C4', -0.5)).toBe('#1F3864')
    expect(resolveStyleColor({ theme: 0 })).toBe('#FFFFFF')
    expect(resolveStyleColor({ theme: 1 })).toBe('#000000')
    expect(resolveStyleColor({ theme: 4, tint: 0.5 })).toBe('#A1B8E1')
  })

  it('falls back deterministically for malformed hex without NaN leaks', () => {
    const malformed = ['ZZZZZZ', '', 'FF0000', 'FFF', '#FFF', '#FF000', '#FF00000', ' #FF0000']
    for (const input of malformed) {
      const first = applyTint(input, 0.5)
      const second = applyTint(input, 0.5)
      expect(first).toBe('#000000')
      expect(second).toBe(first)
      expectCleanHex(first)
      expect(applyTint(input, 0)).toBe('#000000')
      expect(applyTint(input, -0.5)).toBe('#000000')
      expectCleanHex(resolveStyleColor(input))
      expect(resolveStyleColor(input)).toBe('#000000')
    }
  })

  it('treats hash-prefixed red as valid since the file requires #RRGGBB', () => {
    expect(applyTint('#FF0000', 0.5)).not.toBe('#000000')
    expectCleanHex(applyTint('#FF0000', 0.5))
    expect(resolveStyleColor('#FF0000')).toBe('#FF0000')
  })

  it('falls back when a custom palette entry is malformed', () => {
    const resolved = resolveStyleColor({ theme: 0, tint: 0.5 }, ['ZZZZZZ'])
    expect(resolved).toBe('#000000')
    expectCleanHex(resolved)
  })
})
