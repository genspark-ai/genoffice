import { describe, expect, it } from 'vitest'
import { cssFontFamily } from '../src/renderer/text-layout/font-rules'
import { fontAttrsFromFamilyChain } from '../src/renderer/editor/marks'

describe('Avenir Next LT Pro (M365 cloud face) chain', () => {
  it("falls to macOS's Avenir Next, not the generic sans", () => {
    expect(cssFontFamily('Avenir Next LT Pro')).toMatch(/^'Avenir Next LT Pro','Avenir Next',/)
    expect(cssFontFamily('Avenir Next LT Pro Demi')).toMatch(
      /^'Avenir Next LT Pro Demi','Avenir Next Demi Bold','Avenir Next',/,
    )
    expect(cssFontFamily('Avenir Next LT Pro Light')).toMatch(
      /^'Avenir Next LT Pro Light','Avenir Light','Avenir Next',/,
    )
  })

  it('round-trips to the declared face', () => {
    expect(fontAttrsFromFamilyChain(cssFontFamily('Avenir Next LT Pro'))).toEqual({
      font: 'Avenir Next LT Pro',
      fontAscii: 'Avenir Next LT Pro',
    })
  })
})
