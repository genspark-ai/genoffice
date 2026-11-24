/**
 * Word truth (probe 2026-09-01, MS Mincho 10.5pt, w:jc="both", both
 * characterSpacingControl values): compression exists only under
 * compressPunctuation; a full line's 、。/closing brackets shrink by a uniform
 * per-glyph amount sized to the deficit; a kinsoku pull (next chars cannot
 * start a line) compresses to about half width, a voluntary pull only lightly;
 * the ragged last line never compresses.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  mapTextNodes,
  decideCjkHang,
  decideCjkHangUnderCompression,
  decideCjkShrinks,
  decideKanaShrink,
  forbidsLineEnd,
  forbidsLineStart,
  shrinkTargets,
  isCompressible,
  isCompressibleKana,
  measuresAlignment,
  setSpacingControl,
  spacingControlOf,
  shrinkStyle,
  spacingHost,
  usesEastAsianRules,
  type CjkPunctShrinkStorage,
  type HangLineChars,
  type KanaLineChars,
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
describe('decideCjkHangUnderCompression', () => {
  // Yu Mincho 10.5pt probes 2026-10-13 scaled to EM = 14: a kana + stop chain
  // whose stop cannot fit even at half width
  const chain = (over: Partial<ShrinkLineChars>) =>
    line({
      punctCount: 1,
      candWidths: [EM, EM],
      candPunctCount: 1,
      candEndsWithStop: true,
      forced: true,
      ...over,
    })

  it('compresses the own closer by the body deficit and hangs the stop', () => {
    // body = 604 + 14 = 618 vs 616: 2 (+eps) over the one comma
    expect(decideCjkHangUnderCompression(chain({ natural: 604 }))).toBeCloseTo(2.5, 5)
  })

  it('hangs without compression when the body already fits', () => {
    expect(decideCjkHangUnderCompression(chain({ natural: 600 }))).toBe(0)
    expect(decideCjkHangUnderCompression(chain({ natural: 600, punctCount: 0 }))).toBe(0)
  })

  it('declines past half a glyph in total, however many closers share it', () => {
    expect(decideCjkHangUnderCompression(chain({ natural: 609.5 }))).toBeNull()
    expect(decideCjkHangUnderCompression(chain({ natural: 609.5, punctCount: 2 }))).toBeNull()
  })

  it('needs an own closer for a deficit and a chain ending with a stop', () => {
    expect(decideCjkHangUnderCompression(chain({ natural: 604, punctCount: 0 }))).toBeNull()
    expect(
      decideCjkHangUnderCompression(chain({ natural: 604, candEndsWithStop: false })),
    ).toBeNull()
    expect(decideCjkHangUnderCompression(chain({ natural: 604, candWidths: [] }))).toBeNull()
  })

  it('leaves a line with a hung stop to the hang rule', () => {
    expect(
      decideCjkShrinks([chain({ natural: 632, punctCount: 3, candWidths: [], hungWidth: EM })]),
    ).toEqual([null])
  })

  it('keeps a hung stop: the body excludes it, the pool excludes it', () => {
    expect(
      decideCjkHangUnderCompression(
        chain({ natural: 632, punctCount: 2, candWidths: [], hungWidth: EM }),
      ),
    ).toBeCloseTo(2.5, 5)
    expect(
      decideCjkHangUnderCompression(
        chain({ natural: 640, punctCount: 2, candWidths: [], hungWidth: EM }),
      ),
    ).toBeNull()
  })
})

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

type FakeChild = {
  isText: boolean
  text?: string
  marks: unknown[]
  type: { name: string }
  nodeSize: number
}
function fakePara(children: FakeChild[]): ProseMirrorNode {
  return {
    forEach(cb: (child: FakeChild, offset: number) => void) {
      let offset = 0
      for (const c of children) {
        cb(c, offset)
        offset += c.nodeSize
      }
    },
  } as unknown as ProseMirrorNode
}
const text = (t: string): FakeChild => ({
  isText: true,
  text: t,
  marks: [],
  type: { name: 'text' },
  nodeSize: t.length,
})
const hardBreak: FakeChild = { isText: false, marks: [], type: { name: 'hardBreak' }, nodeSize: 1 }
const image: FakeChild = { isText: false, marks: [], type: { name: 'docImage' }, nodeSize: 1 }

describe('mapTextNodes', () => {
  const ea = () => true
  it('maps decoration-split DOM text nodes back to document positions', () => {
    const p = document.createElement('p')
    p.innerHTML = 'ab<span class="x">c</span>d<br>ef'
    const node = fakePara([text('abcd'), hardBreak, text('ef')])
    const mapped = mapTextNodes(p, node, 100, ea)
    expect(mapped).not.toBeNull()
    expect(mapped!.slots.map((s) => [s.dom.data, s.from])).toEqual([
      ['ab', 101],
      ['c', 103],
      ['d', 104],
      ['ef', 106],
    ])
    expect(mapped!.breaks).toEqual([105])
  })

  it('refuses text the document does not have (widgets, cursor wrappers) and atoms', () => {
    const p = document.createElement('p')
    p.innerHTML = 'ab<span contenteditable="false">[1]</span>cd'
    expect(mapTextNodes(p, fakePara([text('abcd')]), 0, ea)).toBeNull()
    p.innerHTML = 'ab\ufeffcd'
    expect(mapTextNodes(p, fakePara([text('abcd')]), 0, ea)).toBeNull()
    p.innerHTML = 'abc'
    expect(mapTextNodes(p, fakePara([text('abcd')]), 0, ea)).toBeNull()
    p.innerHTML = 'ab<img>cd'
    expect(mapTextNodes(p, fakePara([text('ab'), image, text('cd')]), 0, ea)).toBeNull()
  })
})

describe('spacingHost', () => {
  it('reads the base letter-spacing past an active shrink span, so a second pass sees the same base', () => {
    const p = document.createElement('p')
    p.style.letterSpacing = '2px'
    p.innerHTML =
      '\u6f22<span class="doc-east-asian-font" style="letter-spacing:4px"><span class="doc-cjkshrink" style="letter-spacing:-3px">\u3002</span></span><span class="doc-cjkshrink" style="letter-spacing:-1px">\u3001</span>'
    document.body.append(p)
    const baseOf = (el: Element | null) =>
      parseFloat(getComputedStyle(spacingHost(el)!).letterSpacing)
    const [ea, bare] = [...p.querySelectorAll('.doc-cjkshrink')]
    expect(spacingHost(ea.firstChild!.parentElement)).toBe(ea.parentElement)
    expect(baseOf(ea)).toBe(4)
    expect(baseOf(bare)).toBe(2)
    expect(spacingHost(p)).toBe(p)
    // pass two: the shrink decoration changes but the base it is computed from does not
    ;(ea as HTMLElement).style.letterSpacing = '-5px'
    expect(baseOf(ea)).toBe(4)
    p.remove()
  })
})

describe('native fullwidth-closer trimming', () => {
  it('is off on every text surface so a line-final closer wraps with its predecessor', () => {
    const css = readFileSync(resolve('src/renderer/styles.css'), 'utf8')
    const rule = /\n((?:\.[\w-]+,\n)*\.[\w-]+) \{\n\s*text-spacing-trim: space-all;\n\}/.exec(css)
    const roots = rule?.[1].split(',\n') ?? []
    expect(roots).toEqual(
      expect.arrayContaining(['.doc-page', '.pv-content', '.page-hf', '.doc-textbox-editor']),
    )
  })
})

/**
 * Word truth (probes 2026-11-18 and 2026-11-22, compressPunctuationAndJapaneseKana,
 * w:jc="both", 12pt in a 468pt column unless said): the numbers are the
 * advances Word's PDF shows; ours add the half-pixel overshoot that makes
 * Chromium take the run up.
 */
describe('decideKanaShrink', () => {
  const EPS = 0.5
  const STOP = { width: 24, kanaWidth: 12, kanaCount: 1, openerCredit: 0, stopWidth: 12 }
  const ITER = { ...STOP, stopWidth: 0 }
  const ONE = { width: 12, kanaWidth: 12, kanaCount: 1, openerCredit: 0, stopWidth: 0 }
  /** `kana` kana among 38 characters, the 39th and the mark after it on the next line */
  function line(kana: number, over: Partial<KanaLineChars> = {}): KanaLineChars {
    return {
      natural: 456,
      avail: 468,
      chars: 38,
      kanaWidth: 12 * kana,
      kanaCount: kana,
      punctCredit: 0,
      chain: STOP,
      hangs: true,
      ...over,
    }
  }
  const kanaOnly = (need: number, count: number, halfStop = false) => ({
    perKana: (need + EPS) / count,
    halfStop,
    punct: 0,
  })

  it('brings a stop inside when the kana can give its whole advance (16 kana: 11.25pt each)', () => {
    expect(decideKanaShrink(line(15))).toEqual(kanaOnly(12, 16))
    expect(decideKanaShrink(line(38))).toEqual(kanaOnly(12, 39))
  })

  it('sets the stop half width when they can give half of it (8 to 15 kana)', () => {
    expect(decideKanaShrink(line(14))).toEqual(kanaOnly(6, 15, true))
    expect(decideKanaShrink(line(7))).toEqual(kanaOnly(6, 8, true))
  })

  it('leaves the stop to hang below that (7 kana)', () => {
    expect(decideKanaShrink(line(6))).toBeNull()
    expect(decideKanaShrink(line(0))).toBeNull()
  })

  it('takes a mark that cannot hang up whole or not at all (16 kana yes, 15 no)', () => {
    expect(decideKanaShrink(line(15, { chain: ITER }))).toEqual(kanaOnly(12, 16))
    expect(decideKanaShrink(line(14, { chain: ITER }))).toBeNull()
  })

  it('caps the total at a sixteenth of the kana advances', () => {
    // 39 kana + two iteration marks: 24pt of 29.25 (11.385pt each); three marks need 36
    const two = { ...ITER, width: 36 }
    expect(decideKanaShrink(line(38, { chain: two }))?.perKana).toBeCloseTo((24 + EPS) / 39, 5)
    expect(decideKanaShrink(line(38, { chain: { ...ITER, width: 48 } }))).toBeNull()
    // an ASCII bracket is 6pt wide: 39 kana at 11.85pt
    const paren = { ...ITER, width: 18 }
    expect(decideKanaShrink(line(38, { chain: paren }))?.perKana).toBeCloseTo((6 + EPS) / 39, 5)
  })

  it('with w:overflowPunct off the stop is half width or stays down', () => {
    expect(decideKanaShrink(line(38, { hangs: false }))).toEqual(kanaOnly(6, 39, true))
    expect(decideKanaShrink(line(9, { hangs: false }))?.halfStop).toBe(true)
    expect(decideKanaShrink(line(3, { hangs: false }))).toBeNull()
  })

  it('measures the cap on the grid advance (13pt pitch: 8 kana at 12.19pt)', () => {
    const grid = { width: 26, kanaWidth: 13, kanaCount: 1, openerCredit: 0, stopWidth: 13 }
    const on = (kana: number) =>
      line(kana, { natural: 455, chars: 35, kanaWidth: 13 * kana, chain: grid })
    expect(decideKanaShrink(on(7))?.perKana).toBeCloseTo((6.5 + EPS) / 8, 5)
    expect(decideKanaShrink(on(6))).toBeNull()
  })

  it('an opening bracket taken up gives its blank half first (18 kana at 11.67pt)', () => {
    const opener = { width: 24, kanaWidth: 0, kanaCount: 0, openerCredit: 6, stopWidth: 0 }
    expect(decideKanaShrink(line(18, { chain: opener }))?.perKana).toBeCloseTo((6 + EPS) / 18, 5)
  })

  it('a bracket that fits once its blank half is gone comes up without any kana giving (10.5pt)', () => {
    // 43 characters, 16.5pt left over; bracket + kana need 21, the bracket gives 5.25
    const run = { width: 21, kanaWidth: 10.5, kanaCount: 1, openerCredit: 5.25, stopWidth: 0 }
    const on = (kana: number) =>
      line(kana, { natural: 451.5, chars: 43, kanaWidth: 10.5 * kana, chain: run })
    expect(decideKanaShrink(on(43))).toEqual({ perKana: 0, halfStop: false, punct: 0 })
    // kanji only, and the character behind the bracket a kanji too: nothing has to carry an overshoot
    const kanji = { ...run, kanaWidth: 0, kanaCount: 0 }
    expect(decideKanaShrink({ ...on(0), chain: kanji })).toEqual({
      perKana: 0,
      halfStop: false,
      punct: 0,
    })
    // a run that fits as it is was never pushed down: nothing to do
    expect(
      decideKanaShrink({ ...on(43), chain: { ...run, openerCredit: 0, width: 15 } }),
    ).toBeNull()
  })

  it('takes the next kana up when the line would stretch enough otherwise (50 kana, 1-11pt left over)', () => {
    const left = (pt: number, kana = 38) => line(kana, { avail: 456 + pt, chain: ONE })
    // 39 kana: from 3pt left over (11.78pt each), not with 2
    expect(decideKanaShrink(left(3))).toEqual(kanaOnly(9, 39))
    expect(decideKanaShrink(left(11))).toEqual(kanaOnly(1, 39))
    expect(decideKanaShrink(left(2))).toBeNull()
    expect(decideKanaShrink(left(1))).toBeNull()
    // 20 kanji + 19 kana: from 4pt left over (11.60pt each), not with 3
    expect(decideKanaShrink(left(4, 18))).toEqual(kanaOnly(8, 19))
    expect(decideKanaShrink(left(3, 18))).toBeNull()
    // 9pt left over: the cap alone decides (4 kana at 11.25pt, 3 cannot)
    expect(decideKanaShrink(left(9, 3))).toEqual(kanaOnly(3, 4))
    expect(decideKanaShrink(left(9, 2))).toBeNull()
  })

  it('a line filled to the margin takes nothing up (50 kana at 12pt, 39 to the line)', () => {
    expect(decideKanaShrink(line(39, { natural: 468, chars: 39, chain: ONE }))).toBeNull()
    expect(decideKanaShrink(line(39, { natural: 468, chars: 39, chain: STOP }))).toBeNull()
  })

  it('10.5pt, 6pt left over after 44: the 45th kana comes up (10.4pt), not after 44 kanji', () => {
    const one = { ...ONE, width: 10.5, kanaWidth: 10.5 }
    const d = decideKanaShrink(line(44, { natural: 462, chars: 44, kanaWidth: 462, chain: one }))
    expect(d?.perKana).toBeCloseTo((4.5 + EPS) / 45, 5)
    expect(
      decideKanaShrink(line(0, { natural: 462, chars: 44, kanaWidth: 0, chain: one })),
    ).toBeNull()
    // 0.75pt left over behind a bracket: the next kana stays down
    const tight = line(44, { natural: 467.25, chars: 45, kanaWidth: 462, chain: one })
    expect(decideKanaShrink(tight)).toBeNull()
  })

  it('closing punctuation gives half its advance first, the kana the rest (one comma + 8 kana)', () => {
    const comma = { punctCredit: 6, chain: ITER }
    expect(decideKanaShrink(line(7, comma))).toEqual({
      perKana: (6 + EPS) / 8,
      halfStop: false,
      punct: 1,
    })
    expect(decideKanaShrink(line(5, comma))).toBeNull()
    // two marks behind one comma and 24 kana: 6 + 18 (11.25pt each)
    expect(
      decideKanaShrink(line(23, { punctCredit: 6, chain: { ...ITER, width: 36 } }))?.perKana,
    ).toBeCloseTo((18 + EPS) / 24, 5)
    // two commas cover the mark on their own: the kana only carry the overshoot
    expect(decideKanaShrink(line(4, { punctCredit: 12, chain: ITER }))).toEqual({
      perKana: EPS / 5,
      halfStop: false,
      punct: 1,
    })
    expect(decideKanaShrink(line(38, { punctCredit: 12, chain: { ...ITER, width: 18 } }))).toEqual({
      perKana: EPS / 39,
      halfStop: false,
      punct: 0.5,
    })
  })

  it('a line that holds more than fits is not for this rule: the caller sets the run aside first', () => {
    expect(decideKanaShrink(line(38, { natural: 480 }))).toBeNull()
  })
})

describe('the kana class', () => {
  it('is hiragana, katakana and the fullwidth digits and letters', () => {
    for (const ch of '\u3041\u3042\u3083\u3093\u30a1\u30a2\u30f4\u30f5\u30f6\uff10\uff19\uff21\uff5a')
      expect(isCompressibleKana(ch), ch.codePointAt(0)!.toString(16)).toBe(true)
  })

  it('leaves out kanji, marks, halfwidth kana and symbols', () => {
    for (const ch of '\u56fd\u30fc\u3005\u309d\u30fb\u3094\u30f7\u30a0\uff71\uff03\uffe5\u2460\uac00\u3001\u3002A1')
      expect(isCompressibleKana(ch), ch.codePointAt(0)!.toString(16)).toBe(false)
  })
})

describe('characterSpacingControl', () => {
  it('maps the parsed setting to the three radio values', () => {
    expect(spacingControlOf(null)).toBe('none')
    expect(spacingControlOf({})).toBe('none')
    expect(spacingControlOf({ compressPunctuation: true })).toBe('punct')
    expect(spacingControlOf({ compressPunctuation: true, compressKana: true })).toBe('kana')
  })

  it('kana is on only for the kana value; the other two set what they always did', () => {
    const storage = { enabled: false, hangPunct: false, kana: false } as CjkPunctShrinkStorage
    setSpacingControl(storage, 'kana')
    expect(storage).toMatchObject({ enabled: true, hangPunct: false, kana: true })
    setSpacingControl(storage, 'punct')
    expect(storage).toMatchObject({ enabled: true, hangPunct: false, kana: false })
    setSpacingControl(storage, 'none')
    expect(storage).toMatchObject({ enabled: false, hangPunct: true, kana: false })
  })
})
