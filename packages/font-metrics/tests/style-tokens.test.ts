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
})
