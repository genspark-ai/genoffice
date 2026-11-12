import { describe, expect, it } from 'vitest'
import { cssFontFamily, isBundledFont } from '../src/renderer/line-metrics'
import { fontAttrsFromFamilyChain } from '../src/renderer/editor/marks'

describe('Cambria Math font chain', () => {
  it('leads with the Cambria-metric text alias before the math face', () => {
    expect(cssFontFamily('Cambria Math')).toMatch(
      /^'Cambria Math','Cambria Math GO','STIX Two Math','Caladea',/,
    )
    expect(isBundledFont('Cambria Math GO')).toBe(true)
  })

  it('keeps real math faces on STIX Two Math first', () => {
    expect(cssFontFamily('STIX Two Math')).toMatch(/^'STIX Two Math','Caladea',/)
    expect(cssFontFamily('Latin Modern Math')).toMatch(/^'Latin Modern Math','STIX Two Math',/)
  })

  it('round-trips to the declared face, not the alias', () => {
    expect(fontAttrsFromFamilyChain(cssFontFamily('Cambria Math'))).toEqual({
      font: 'Cambria Math',
      fontAscii: 'Cambria Math',
    })
  })
})
