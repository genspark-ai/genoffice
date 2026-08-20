import { describe, expect, it } from 'vitest'
import { buildSpans, groupIntoWords } from '../src/analyze'
import type { PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

const spansOf = (chars: PdfChar[]) => buildSpans(groupIntoWords(chars))

describe('span building', () => {
  it('merges same-styled latin words into one span with literal spaces', () => {
    const { chars } = mkText('hello world', 72)
    const spans = spansOf(chars)
    expect(spans).toHaveLength(1)
    expect(spans[0]!.text).toBe('hello world')
    expect(spans[0]!.script).toBe('latin')
    expect(spans[0]!.dir).toBe('ltr')
  })

  it('splits mixed zh/en text into alternating cjk/latin spans', () => {
    const zh1 = mkText('中文', 72)
    const en = mkText('English', zh1.endX)
    const zh2 = mkText('混排', en.endX)
    const spans = spansOf([...zh1.chars, ...en.chars, ...zh2.chars])
    expect(spans.map((s) => s.text)).toEqual(['中文', 'English', '混排'])
    expect(spans.map((s) => s.script)).toEqual(['cjk', 'latin', 'cjk'])
  })

  it('invisible chars (PDF Tr 3) form their own span carrying the flag (P20)', () => {
    const vis = mkText('正文', 72)
    const hidden = mkText('分节符', vis.endX).chars.map((c) => ({ ...c, invisible: true as const }))
    const spans = spansOf([...vis.chars, ...hidden])
    expect(spans.map((s) => s.text)).toEqual(['正文', '分节符'])
    expect(spans[0]!.invisible).toBeUndefined()
    expect(spans[1]!.invisible).toBe(true)
  })

  it('attaches common punctuation/digits to the surrounding script span', () => {
    const { chars } = mkText('version 2.0, final', 72)
    const spans = spansOf(chars)
    expect(spans).toHaveLength(1)
    expect(spans[0]!.text).toBe('version 2.0, final')
  })

  it('splits on bold style change', () => {
    const normal = mkText('normal ', 72)
    const bold = mkText('bold', normal.endX, { fontWeight: 700 })
    const spans = spansOf([...normal.chars, ...bold.chars])
    expect(spans.map((s) => s.text)).toEqual(['normal ', 'bold'])
    expect(spans[0]!.bold).toBe(false)
    expect(spans[1]!.bold).toBe(true)
  })

  it('detects bold from the font name when the weight is missing', () => {
    const { chars } = mkText('head', 72, { fontFamily: 'Arial-BoldMT' })
    expect(spansOf(chars)[0]!.bold).toBe(true)
  })

  it('measures glyph scale and negative tracking of squeezed text (P5)', () => {
    // 10pt CJK squeezed to 90% glyph width (loose = 9pt) with −0.4pt tracking
    const chars = []
    let x = 72
    for (const ch of '压缩的文本内容') {
      const c = mkChar(ch, x, { fontSize: 10, width: 9 })
      c.hscale = 0.9
      chars.push(c)
      x += 8.6
    }
    const [span] = spansOf(chars)
    expect(span!.charScale).toBeCloseTo(0.9, 5)
    expect(span!.charSpacingPt).toBeCloseTo(-0.4, 5)
  })

  it('drops negative tracking when word spaces render at normal width (P14 B: inflated /Widths)', () => {
    // PowerPoint-export pattern: every glyph's declared advance (loose box) is
    // ~0.18em wider than the TJ-laid advance, so pairwise "tracking" reads a
    // large negative constant — but the spaces are normal-width, so the text
    // is not visually squeezed and no w:spacing must be emitted
    const chars: PdfChar[] = []
    let x = 100
    for (const ch of 'Why Dual') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize: 14, width: 4.2 })
        chars.push(sp)
        x += 4.2
        continue
      }
      const c = mkChar(ch, x, { fontSize: 14, width: 7 })
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 10 } // inflated nominal
      chars.push(c)
      x += 7.5 // actual advance: 2.5pt narrower than nominal
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('Why Dual')
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('drops extreme negative tracking when italic overhang shaves the space gap (P15 A)', () => {
    // HTML-export deck pattern: /Widths inflated ~0.24em per glyph AND an
    // italic serif whose ink boxes overlap neighbors (negative intra-word ink
    // gaps). The space ink gap measures just under the 0.18em healthy bar,
    // but adding the overhang back proves the words are normally spaced.
    const chars: PdfChar[] = []
    let x = 100
    const fontSize = 14
    for (const ch of 'Street Food') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize, width: 3.2 })
        chars.push(sp)
        x += 3.2 // ink gap to next char: 3.2 − 0.9 overhang ≈ 0.164em < 0.18em
        continue
      }
      const c = mkChar(ch, x, { fontSize, width: 7 })
      // ink box overhangs the advance by 0.9pt (italic slant)
      c.box = { ...c.box, x1: c.box.x0 + 7.9 }
      // declared advance inflated: nominal 10.4 vs actual 7 → tracking −3.4pt (−0.243em)
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 10.4 }
      chars.push(c)
      x += 7
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('Street Food')
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('drops extreme negative tracking when the real space chars RENDER at word width (P15 A)', () => {
    // short line, noisy ink gaps (italic overhang eats them unevenly) — but
    // the real space chars' own origin-to-origin advance is a normal 0.23em,
    // proving the words are not squeezed
    const chars: PdfChar[] = []
    let x = 100
    const fontSize = 14
    for (const ch of 'of a stove') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize, width: 3.2 })
        chars.push(sp)
        x += 3.2 // real space advance: 3.2pt = 0.23em (healthy)
        continue
      }
      const c = mkChar(ch, x, { fontSize, width: 7 })
      // heavy uneven overhang: ink runs 2.9pt past the advance, so ink gaps
      // across spaces measure near zero or negative
      c.box = { ...c.box, x1: c.box.x0 + 9.9 }
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 10.4 } // inflated /Widths
      chars.push(c)
      x += 7 // tracking −3.4pt (−0.243em)
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('of a stove')
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('keeps extreme negative tracking when the real spaces render squeezed too', () => {
    // genuine hard compression: the space chars themselves render at 0.06em —
    // neither the advance evidence nor the ink evidence may fire
    const chars: PdfChar[] = []
    let x = 100
    const fontSize = 14
    for (const ch of 'of a stove') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize, width: 0.8 })
        chars.push(sp)
        x += 0.8
        continue
      }
      const c = mkChar(ch, x, { fontSize, width: 7 })
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 10.4 }
      chars.push(c)
      x += 7
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('of a stove')
    expect(span!.charSpacingPt).toBeCloseTo(-3.4, 5)
  })

  it('keeps MILD negative tracking under the italic-overhang correction (genuine tightening)', () => {
    // same overlapping ink boxes, but tracking is only −1pt (−0.07em) — far
    // above the −0.15em artifact bar, so the corrected space gap must NOT
    // erase a plausible genuine compression
    const chars: PdfChar[] = []
    let x = 100
    const fontSize = 14
    for (const ch of 'Street Food') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize, width: 2.2 })
        chars.push(sp)
        x += 2.2 // squeezed space: ink gap ≈ 0.09em, corrected ≈ 0.16em — both under bar anyway
        continue
      }
      const c = mkChar(ch, x, { fontSize, width: 7 })
      c.box = { ...c.box, x1: c.box.x0 + 7.9 }
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 8 }
      chars.push(c)
      x += 7
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('Street Food')
    expect(span!.charSpacingPt).toBeCloseTo(-1, 5)
  })

  it('keeps negative tracking when the spaces are squeezed too (genuine compression)', () => {
    const chars: PdfChar[] = []
    let x = 100
    for (const ch of 'Why Dual') {
      if (ch === ' ') {
        const sp = mkChar(' ', x, { fontSize: 14, width: 1.2 })
        chars.push(sp)
        x += 1.2 // squeezed space: ink gap well under 0.18em
        continue
      }
      const c = mkChar(ch, x, { fontSize: 14, width: 7 })
      c.looseBox = { ...c.looseBox, x1: c.looseBox.x0 + 8 }
      chars.push(c)
      x += 7 // 1pt under nominal on every pair
    }
    const [span] = spansOf(chars)
    expect(span!.text).toBe('Why Dual')
    expect(span!.charSpacingPt).toBeCloseTo(-1, 5)
  })

  it('leaves unsqueezed text without compression fields', () => {
    const { chars } = mkText('normal text here', 72)
    const [span] = spansOf(chars)
    expect(span!.charScale).toBeUndefined()
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('restores POSITIVE letter-spacing from stable wide gaps (P7 B: w:spacing)', () => {
    // 10pt latin letters advancing 5pt (loose) + 1pt tracking each — the
    // gap is stable and far above the 0.05pt noise floor
    const { chars } = mkText('SPACED', 72, { fontSize: 10, tracking: 1 })
    const [span] = spansOf(chars)
    expect(span!.text).toBe('SPACED')
    expect(span!.charSpacingPt).toBeCloseTo(1, 5)
  })

  it('CJK with its natural full-width advance gets no letter-spacing (P7 B guard)', () => {
    // fullwidth glyphs: loose box = advance = 1em, zero extra gap
    const { chars } = mkText('中文标题文字', 72, { fontSize: 12 })
    const [span] = spansOf(chars)
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('positive CJK letter-spacing IS restored when the gap is real', () => {
    const chars = []
    let x = 72
    for (const ch of '字距标题') {
      chars.push(mkChar(ch, x, { fontSize: 12, width: 12 }))
      x += 13.2 // 1.2pt tracking on a 12pt full-width advance
    }
    const [span] = spansOf(chars)
    expect(span!.charSpacingPt).toBeCloseTo(1.2, 5)
  })

  it('a word gap never enters the tracking chain (inferred space resets it)', () => {
    // 5pt inter-word gap (inferred space); intra-word gaps are tight — the
    // wide gap must not read as +5pt letter-spacing
    const a = mkText('AB', 72, { fontSize: 10 })
    const b = mkText('CD', a.endX + 5, { fontSize: 10 })
    const [span] = spansOf([...a.chars, ...b.chars])
    expect(span!.text).toBe('AB CD')
    expect(span!.charSpacingPt).toBeUndefined()
  })

  it('splits on font size change beyond tolerance, keeps within-tolerance together', () => {
    const a = mkText('ab', 72, { fontSize: 10 })
    const b = mkText('cd', a.endX, { fontSize: 10.05 }) // within 0.1pt tolerance
    const c = mkText('ef', b.endX, { fontSize: 14 })
    const spans = spansOf([...a.chars, ...b.chars, ...c.chars])
    expect(spans.map((s) => s.text)).toEqual(['abcd', 'ef'])
  })

  it('splits on color change', () => {
    const black = mkText('black', 72)
    const red = mkText('red', black.endX, { color: 'FF0000' })
    const spans = spansOf([...black.chars, ...red.chars])
    expect(spans.map((s) => s.color)).toEqual(['000000', 'FF0000'])
  })

  it('marks arabic spans rtl (P1 pages degrade, but the field is authoritative)', () => {
    const a = mkChar('م', 72, { width: 5 })
    const b = mkChar('ر', 77, { width: 5 })
    const spans = spansOf([a, b])
    expect(spans[0]!.dir).toBe('rtl')
    expect(spans[0]!.script).toBe('arabic')
  })

  it('keeps fullwidth CJK punctuation inside the cjk span', () => {
    const zh = mkText('你好,世界。', 72)
    const spans = spansOf(zh.chars)
    expect(spans).toHaveLength(1)
    expect(spans[0]!.text).toBe('你好,世界。')
    expect(spans[0]!.script).toBe('cjk')
  })
})

describe('calibrateFontSizes (P8)', () => {
  const loosen = (chars: PdfChar[], ratio: number): PdfChar[] =>
    chars.map((c) => ({
      ...c,
      looseBox: {
        ...c.looseBox,
        y0: c.originY - 0.2 * ratio * c.fontSize,
        y1: c.originY + 0.8 * ratio * c.fontSize,
      },
    }))

  it('replaces a bogus declared size with the metric-derived effective size', async () => {
    const { calibrateFontSizes } = await import('../src/analyze/chars')
    // AI-generated pattern: 15pt declared, loose (advance) box only 0.65 em
    const chars = loosen(mkText('声明字号偏大的正文', 72, { fontSize: 15 }).chars, 0.65)
    calibrateFontSizes(chars)
    for (const c of chars) expect(c.fontSize).toBeCloseTo(15 * 0.65, 5)
  })

  it('keeps declared sizes for normal fonts and degenerate metrics', async () => {
    const { calibrateFontSizes } = await import('../src/analyze/chars')
    // mkChar default loose box is ~1.2 em (normal font)
    const normal = mkText('normal text', 72, { fontSize: 12 }).chars
    // degenerate: loose box collapsed to nothing (broken metrics)
    const broken = loosen(mkText('破损度量', 300, { fontSize: 14 }).chars, 0.01)
    calibrateFontSizes([...normal, ...broken])
    for (const c of normal) expect(c.fontSize).toBe(12)
    for (const c of broken) expect(c.fontSize).toBe(14)
  })
})
