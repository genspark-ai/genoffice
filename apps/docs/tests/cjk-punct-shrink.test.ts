/**
 * Word truth (probe 2026-09-01, MS Mincho 10.5pt, w:jc="both", both
 * characterSpacingControl values): compression exists only under
 * compressPunctuation; a full line's 、。/closing brackets shrink by a uniform
 * per-glyph amount sized to the deficit; a kinsoku pull (next chars cannot
 * start a line) compresses to about half width, a voluntary pull only lightly;
 * the ragged last line never compresses.
 */
import { describe, expect, it } from 'vitest'
import { decideCjkShrinks, type ShrinkLineChars } from '../src/renderer/editor/cjk-punct-shrink'

const EM = 14

function line(over: Partial<ShrinkLineChars>): ShrinkLineChars {
  return {
    natural: 610,
    avail: 616,
    punctCount: 2,
    avgPunctW: EM,
    candWidths: [EM],
    candPunctCount: 0,
    forced: false,
    ...over,
  }
}

describe('decideCjkShrinks', () => {
  it('pulls voluntarily when the deficit stays within a quarter width per glyph', () => {
    const [d] = decideCjkShrinks([line({ natural: 608 })])
    // deficit = (608 - 616) + 14 = 6 (+eps), spread over 2 glyphs
    expect(d).toBeCloseTo(6.5 / 2, 5)
  })

  it('declines a voluntary pull past the quarter-width cap', () => {
    expect(decideCjkShrinks([line({ natural: 616 })])).toEqual([null])
  })

  it('accepts a deeper deficit when the chain is kinsoku-forced', () => {
    const [d] = decideCjkShrinks([
      line({ natural: 604, candWidths: [EM, EM], candPunctCount: 1, forced: true }),
    ])
    // deficit = (604 - 616) + 28 = 16 (+eps); admissible over the post-pull
    // pool of 3, emitted over the 2 glyphs that exist now
    expect(d).toBeCloseTo(16.5 / 2, 5)
  })

  it('declines a forced pull past the half-width floor', () => {
    expect(
      decideCjkShrinks([line({ natural: 616, punctCount: 1, candWidths: [EM, EM], forced: true })]),
    ).toEqual([null])
  })

  it('keeps compression for a line already holding pulled characters', () => {
    const [d] = decideCjkShrinks([line({ natural: 622, candWidths: [] })])
    expect(d).toBeCloseTo(6.5 / 2, 5)
  })

  it('never compresses without a pull candidate (ragged last line)', () => {
    expect(decideCjkShrinks([line({ natural: 600, candWidths: [] })])).toEqual([null])
  })

  it('needs compressible glyphs on the line', () => {
    expect(decideCjkShrinks([line({ punctCount: 0, avgPunctW: 0 })])).toEqual([null])
  })
})
