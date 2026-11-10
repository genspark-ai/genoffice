import { describe, expect, it } from 'vitest'
import { MAX_LENGTH_EMU, parseEmu, parsePoints, parseTwips } from '../src/renderer/ai/lengths'

describe('AI length parsers', () => {
  it('parse normal lengths unchanged', () => {
    expect(parseEmu('1in')).toBe(914400)
    expect(parseEmu('12pt')).toBe(12 * 12700)
    expect(parseTwips('1in')).toBe(1440)
    expect(parsePoints('72pt')).toBe(72)
  })

  it('reject astronomic lengths that would break layout', () => {
    expect(parseEmu('999999in')).toBeUndefined()
    expect(parseEmu('99999999in')).toBeUndefined()
    expect(parseTwips('100000in')).toBeUndefined()
    expect(parseEmu(1e18)).toBeUndefined()
    expect(parseEmu(NaN)).toBeUndefined()
    expect(MAX_LENGTH_EMU).toBe(32_000_000)
  })
})
