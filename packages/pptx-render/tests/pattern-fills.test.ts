import { describe, expect, it } from 'vitest'
import { patternGrid } from '../src/pattern-fills'

const density = (g: boolean[][]) => g.flat().filter(Boolean).length / 64

describe('patternGrid', () => {
  it('percent presets dither at their stated density', () => {
    for (const [name, pct] of [
      ['pct5', 5],
      ['pct25', 25],
      ['pct50', 50],
      ['pct90', 90],
    ] as const) {
      expect(Math.abs(density(patternGrid(name)) - pct / 100)).toBeLessThan(0.05)
    }
    // monotone in density
    expect(density(patternGrid('pct10'))).toBeLessThan(density(patternGrid('pct40')))
  })

  it('structural presets have the right orientation', () => {
    const horz = patternGrid('ltHorz')
    expect(horz[0]!.every(Boolean)).toBe(true)
    expect(horz[1]!.every((v) => !v)).toBe(true)
    const vert = patternGrid('ltVert')
    expect(vert.every((row) => row[0])).toBe(true)
    expect(vert.every((row) => !row[1])).toBe(true)
    // ltDnDiag: one fg pixel per row marching right
    const dn = patternGrid('ltDnDiag')
    for (let v = 0; v < 8; v++) expect(dn[v]![v]).toBe(true)
  })

  it('unknown presets fall back to a visible pattern', () => {
    const g = patternGrid('nonsense42')
    expect(density(g)).toBeGreaterThan(0.05)
    expect(density(g)).toBeLessThan(0.95)
  })
})
