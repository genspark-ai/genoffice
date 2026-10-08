import { describe, expect, it } from 'vitest'
import { styleScore, styleTokens } from '../src/sfnt'

describe('styleTokens', () => {
  it('keeps Demibold distinct from Bold (#1481, #1482)', () => {
    expect(styleTokens('Demibold')).toEqual(['demibold'])
    expect(styleTokens('SourceSansPro-Demibold')).toEqual(['demibold'])
    expect(styleTokens('Helvetica Neue Demi Bold')).toEqual(['demibold'])
    expect(styleTokens('ITCFranklinGothicStd-Demi')).toEqual(['demibold'])
    // Bold / Semibold / ExtraBold tokenization is unchanged
    expect(styleTokens('Bold')).toEqual(['bold'])
    expect(styleTokens('Semibold')).toEqual(['semibold'])
    expect(styleTokens('Extrabold')).toEqual(['extrabold'])
    expect(styleTokens('BoldItalic')).toEqual(['bold', 'italic'])
  })

  it('stops a Demibold face from tying Bold on a bold-only want', () => {
    const bold: Parameters<typeof styleScore>[0] = { path: 'x', offset: 0, style: 'bold' }
    const demi: Parameters<typeof styleScore>[0] = { path: 'x', offset: 0, style: 'demibold' }
    const want = styleTokens('X-Bold')
    expect(styleScore(bold, want)).toBeGreaterThan(styleScore(demi, want))
  })

  it('falls back from a Demibold want to Bold before Regular', () => {
    type Face = Parameters<typeof styleScore>[0]
    const bold: Face = { path: 'x', offset: 0, style: 'Bold' }
    const regular: Face = { path: 'x', offset: 0, style: 'Regular' }
    const demi: Face = { path: 'x', offset: 0, style: 'Demibold' }
    const want = styleTokens('Foo-Demibold')
    expect(styleScore(demi, want)).toBeGreaterThan(styleScore(bold, want))
    expect(styleScore(bold, want)).toBeGreaterThan(styleScore(regular, want))
    expect(styleScore(demi, styleTokens('Foo-Bold'))).toBeGreaterThan(
      styleScore(regular, styleTokens('Foo-Bold')),
    )
    expect(styleScore(regular, [])).toBeGreaterThan(styleScore(bold, []))
  })
})
