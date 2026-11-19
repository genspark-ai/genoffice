/**
 * `a:alpha` is a percentage. A negative value used to reach `toString(16)`
 * as a negative number, so the colour came out as `#FF0000-27D` — not a
 * colour at all, and every consumer that slices the last two characters
 * read the wrong alpha. Clamp the byte before formatting.
 */
import { describe, it, expect } from 'vitest'
import { applyColorMods } from '../src/color'

const alpha = (value: string) => applyColorMods('#FF0000', { 'a:alpha': { '@_val': value } } as any)

describe('a colour alpha modifier', () => {
  it('still formats a normal value as two hex digits', () => {
    expect(alpha('50000')).toBe('#FF000080')
    expect(alpha('0')).toBe('#FF000000')
    expect(alpha('100000')).toBe('#FF0000') // fully opaque: no suffix
  })

  it('clamps a negative value to transparent instead of writing a negative byte', () => {
    expect(alpha('-250000')).toBe('#FF000000')
    expect(alpha('-100000')).toBe('#FF000000')
  })

  it('never emits a suffix that is not two hex digits', () => {
    for (const value of ['-250000', '-1', '1', '99999', '50000']) {
      expect(alpha(value)).toMatch(/^#[0-9A-F]{6}([0-9A-F]{2})?$/)
    }
  })
})
