/**
 * Word truth (probe 2026-09-01, MS Mincho 10.5pt, w:jc="both", both
 * characterSpacingControl values): compression exists only under
 * compressPunctuation; a full line's 、。/closing brackets shrink by a uniform
 * per-glyph amount sized to the deficit; a kinsoku pull (next chars cannot
 * start a line) compresses to about half width, a voluntary pull only lightly;
 * the ragged last line never compresses.
 */
import { describe, expect, it } from 'vitest'
import {
  decideCjkHang,
  decideCjkShrinks,
  forbidsLineEnd,
  forbidsLineStart,
  shrinkTargets,
  isCompressible,
  measuresAlignment,
  shrinkStyle,
  usesEastAsianRules,
  type HangLineChars,
  type ShrinkLineChars,
} from '../src/renderer/editor/cjk-punct-shrink'

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

  it('spreads an accepted pull over opening brackets without counting them toward the cap', () => {
    const [d] = decideCjkShrinks([line({ natural: 608, openCount: 1 })])
    // admissible over the 2 closing glyphs (3.25 <= 0.27em), emitted over 3
    expect(d).toBeCloseTo(6.5 / 3, 5)
    // 10.5 over 3 glyphs would pass the cap; over the 2 closing glyphs it does not
    expect(decideCjkShrinks([line({ natural: 612, openCount: 1 })])).toEqual([null])
  })

  it('needs compressible glyphs on the line or in a forced chain', () => {
    expect(decideCjkShrinks([line({ punctCount: 0, avgPunctW: 0 })])).toEqual([null])
    // a voluntary pull has nothing to compress
    expect(decideCjkShrinks([line({ natural: 604, punctCount: 0 })])).toEqual([null])
  })

  it('lets a pulled stop absorb the deficit when it is the only compressible glyph', () => {
    // 2026-09-05 probe: MS Mincho 10.5pt, 45 chars + trailing stop in a 482pt
    // column under compressPunctuation: the stop shrinks to half width
    const [d] = decideCjkShrinks([
      line({ natural: 594, punctCount: 0, candWidths: [EM, EM], candPunctCount: 1, forced: true }),
    ])
    // deficit = (594 - 616) + 28 = 6 (+eps) on the stop itself (<= half width)
    expect(d).toBeCloseTo(6.5, 5)
    expect(
      decideCjkShrinks([
        line({
          natural: 596,
          punctCount: 0,
          candWidths: [EM, EM],
          candPunctCount: 1,
          forced: true,
        }),
      ]),
    ).toEqual([null])
  })

  it("lands a pulled-stop amount on the chain, not on the line's opening brackets", () => {
    const own = ['「']
    const cand = ['。']
    const noCloser = line({
      natural: 594,
      punctCount: 0,
      candWidths: [EM, EM],
      candPunctCount: 1,
      forced: true,
    })
    expect(shrinkTargets(noCloser, own, cand)).toBe(cand)
    const withCloser = line({ natural: 594, punctCount: 1, candWidths: [EM], candPunctCount: 0 })
    expect(shrinkTargets(withCloser, own, cand)).toBe(own)
  })
})

/**
 * Word truth (probe 2026-09-05, MS Mincho 10.5pt, 482pt column, compat 15
 * justified, no characterSpacingControl / doNotCompress): 45 chars plus a
 * trailing 。、」）． hang the stop with its origin exactly at the right margin;
 * two characters over, or a left-aligned paragraph, drop the pair instead.
 */
describe('decideCjkHang', () => {
  function hang(over: Partial<HangLineChars>): HangLineChars {
    return {
      natural: 616,
      avail: 643,
      hungWidth: null,
      candWidths: [EM, EM],
      candEndsWithStop: true,
      ...over,
    }
  }

  it('hangs the trailing stop when everything before it fits', () => {
    // 44 chars (616) + 1 body char (14) = 630 <= 643, stop hangs
    expect(decideCjkHang(hang({}))).toBe('pull')
  })

  it('refuses when the body itself does not fit', () => {
    expect(decideCjkHang(hang({ natural: 630 }))).toBeNull()
  })

  it('hangs only stops and closing brackets', () => {
    expect(decideCjkHang(hang({ candEndsWithStop: false }))).toBeNull()
    expect(decideCjkHang(hang({ candWidths: [] }))).toBeNull()
  })

  it('hangs a bare stop chain and a stop after an opening bracket', () => {
    expect(decideCjkHang(hang({ natural: 630, candWidths: [EM] }))).toBe('pull')
    expect(decideCjkHang(hang({ natural: 602, candWidths: [EM, EM, EM] }))).toBe('pull')
  })

  it('keeps a hung stop while the rest of the line fits', () => {
    expect(decideCjkHang(hang({ natural: 644, hungWidth: EM }))).toBe('keep')
    expect(decideCjkHang(hang({ natural: 658, hungWidth: EM }))).toBeNull()
  })
})

describe('compressible glyphs', () => {
  it('covers stops, closing and opening brackets but not ordinary kana', () => {
    for (const ch of '、。」）（「') expect(isCompressible(ch)).toBe(true)
    for (const ch of 'あ漢ー！') expect(isCompressible(ch)).toBe(false)
  })

  it('trims the trailing blank of stops and the leading blank of opening brackets', () => {
    expect(shrinkStyle('。', 2.5, 0)).toBe('letter-spacing:-2.5px')
    expect(shrinkStyle('）', 2.5, 1)).toBe('letter-spacing:-1.5px')
    expect(shrinkStyle('（', 2.5, 1)).toBe('margin-left:-2.5px')
  })

  // Word pulls a trailing ideograph and its ASCII closer onto the line by
  // compressing the comma pool; the chain must carry the closer or the pull is never sized
  it('drags ASCII closers along with a pulled character like JIS ones', () => {
    for (const ch of ')]},.!?:;%』」。、') expect(forbidsLineStart(ch)).toBe(true)
    for (const ch of '(等あa「（') expect(forbidsLineStart(ch)).toBe(false)
    for (const ch of '(「（[') expect(forbidsLineEnd(ch)).toBe(true)
    for (const ch of ')」あ') expect(forbidsLineEnd(ch)).toBe(false)
    expect(isCompressible(')')).toBe(false)
  })
})

describe('measuresAlignment', () => {
  // Word probe 2026-09-04 (Yu Mincho 11pt, Meiryo 14pt, 425pt column): under
  // compatibilityMode 15 a left-aligned line neither pulls nor compresses
  // (39 chars stay 38, 「よ。」 drops to the next line), under compat 14 the
  // same lines pull with compression regardless of alignment
  it('measures only justified paragraphs under the Word 2013+ layout', () => {
    expect(measuresAlignment('justify', false)).toBe(true)
    expect(measuresAlignment('left', false)).toBe(false)
    expect(measuresAlignment('start', false)).toBe(false)
    expect(measuresAlignment('center', false)).toBe(false)
  })

  it('measures every alignment under the legacy layout', () => {
    expect(measuresAlignment('left', true)).toBe(true)
    expect(measuresAlignment('justify', true)).toBe(true)
  })
})

/**
 * Word truth (probe 2026-09-05): hanging, compression and kinsoku follow the
 * run's effective w:lang w:eastAsia — ja-JP / zh-CN (or no w:lang at all)
 * apply them; en-US, en-GB and ko-KR runs get plain wrapping, the next line
 * may even start with the 。.
 */
describe('usesEastAsianRules', () => {
  it('enables the rules for CJK languages and for a missing language', () => {
    expect(usesEastAsianRules('ja-JP')).toBe(true)
    expect(usesEastAsianRules('zh-CN')).toBe(true)
    expect(usesEastAsianRules('zh-TW')).toBe(true)
    expect(usesEastAsianRules('ja')).toBe(true)
    expect(usesEastAsianRules(null)).toBe(true)
    expect(usesEastAsianRules(undefined)).toBe(true)
  })

  it('disables them for non-CJK East Asian languages', () => {
    expect(usesEastAsianRules('en-US')).toBe(false)
    expect(usesEastAsianRules('en-GB')).toBe(false)
    expect(usesEastAsianRules('ko-KR')).toBe(false)
  })
})
