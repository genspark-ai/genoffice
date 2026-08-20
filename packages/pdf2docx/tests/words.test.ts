import { describe, expect, it } from 'vitest'
import { groupIntoWords, spaceGapThreshold } from '../src/analyze'
import { mkChar, mkText } from './helpers/chars'

const texts = (words: ReturnType<typeof groupIntoWords>) => words.map((w) => w.text)

describe('latin word segmentation', () => {
  it('splits on real space characters', () => {
    const { chars } = mkText('ab cd', 72)
    const words = groupIntoWords(chars)
    expect(texts(words)).toEqual(['ab', 'cd'])
    expect(words[0]!.spaceBefore).toBe(false)
    expect(words[1]!.spaceBefore).toBe(true)
  })

  it('infers a space from a wide x gap (no space glyph in the PDF)', () => {
    const a = mkText('ab', 72, { fontSize: 10 })
    const b = mkText('cd', a.endX + 3, { fontSize: 10 }) // 3pt > 0.18 × 10
    const words = groupIntoWords([...a.chars, ...b.chars])
    expect(texts(words)).toEqual(['ab', 'cd'])
    expect(words[1]!.spaceBefore).toBe(true)
  })

  it('keeps tight kerning gaps inside one word', () => {
    const { chars } = mkText('word', 72, { fontSize: 10, tracking: 0.5 }) // 0.5pt << threshold
    expect(texts(groupIntoWords(chars))).toEqual(['word'])
  })

  it('adapts the gap threshold to the font size (same absolute gap, bigger type)', () => {
    // 3pt gap: a word break at 10pt (threshold 1.8pt) but kerning at 24pt (threshold 4.32pt)
    const small = (() => {
      const a = mkText('ab', 72, { fontSize: 10 })
      const b = mkText('cd', a.endX + 3, { fontSize: 10 })
      return groupIntoWords([...a.chars, ...b.chars])
    })()
    const large = (() => {
      const a = mkText('ab', 72, { fontSize: 24 })
      const b = mkText('cd', a.endX + 3, { fontSize: 24 })
      return groupIntoWords([...a.chars, ...b.chars])
    })()
    expect(texts(small)).toEqual(['ab', 'cd'])
    expect(texts(large)).toEqual(['abcd'])
  })

  it('adapts the threshold to loose tracking (median intra-char gap raises the bar)', () => {
    // tracking 1pt at 10pt: base threshold 1.8pt would misread every gap of
    // 1.9pt as a space; the adaptive rule lifts it to median(1) + 1 = 2pt
    const { chars } = mkText('letterspaced', 72, { fontSize: 10, tracking: 1 })
    const tail = mkText('xy', chars[chars.length - 1]!.box.x1 + 1.9, { fontSize: 10 })
    const words = groupIntoWords([...chars, ...tail.chars])
    expect(texts(words)).toEqual(['letterspacedxy'])
  })

  it('spaceGapThreshold scales with font size and tracking', () => {
    expect(spaceGapThreshold(10, 0)).toBeCloseTo(1.8)
    expect(spaceGapThreshold(20, 0)).toBeCloseTo(3.6)
    expect(spaceGapThreshold(10, 1)).toBeCloseTo(2)
    // tight tracking keeps the base
    expect(spaceGapThreshold(10, 0.3)).toBeCloseTo(1.8)
  })
})

describe('CJK / Thai spacing (hard rule: never machine-insert)', () => {
  it('never inserts spaces between CJK chars regardless of gaps', () => {
    // fullwidth glyphs with an extra 5pt gap — way past any Latin threshold
    const a = mkChar('中', 72, { fontSize: 10 })
    const b = mkChar('文', 87, { fontSize: 10 })
    const c = mkChar('字', 102, { fontSize: 10 })
    const words = groupIntoWords([a, b, c])
    expect(texts(words)).toEqual(['中', '文', '字'])
    expect(words.every((w) => !w.spaceBefore)).toBe(true)
  })

  it('never inserts a space at a CJK↔Latin boundary', () => {
    const zh = mkChar('中', 72, { fontSize: 10 })
    const latin = mkText('abc', 87, { fontSize: 10 }) // 5pt gap
    const words = groupIntoWords([zh, ...latin.chars])
    expect(texts(words)).toEqual(['中', 'abc'])
    expect(words[1]!.spaceBefore).toBe(false)
  })

  it('keeps author-written spaces between CJK chars', () => {
    const a = mkChar('中', 72)
    const sp = mkChar(' ', 82, { width: 2.5 })
    const b = mkChar('文', 84.5)
    const words = groupIntoWords([a, sp, b])
    expect(texts(words)).toEqual(['中', '文'])
    expect(words[1]!.spaceBefore).toBe(true)
  })

  it('never inserts spaces between Thai chars', () => {
    const a = mkChar('ก', 72, { fontSize: 10, width: 5 })
    const b = mkChar('ข', 80, { fontSize: 10, width: 5 }) // 3pt gap
    const words = groupIntoWords([a, b])
    expect(texts(words)).toEqual(['ก', 'ข'])
    expect(words.every((w) => !w.spaceBefore)).toBe(true)
  })

  it('kana chars each form their own word without spaces', () => {
    const a = mkChar('こ', 72)
    const b = mkChar('ん', 84) // 2pt gap
    expect(texts(groupIntoWords([a, b]))).toEqual(['こ', 'ん'])
  })

  it('hangul follows the Latin rule (Korean writes real spaces)', () => {
    const { chars } = mkText('안녕 세계', 72, { fontSize: 10 })
    const words = groupIntoWords(chars)
    expect(texts(words)).toEqual(['안녕', '세계'])
    expect(words[1]!.spaceBefore).toBe(true)
  })

  it('drops PDFium-generated spaces between CJK chars (P7: Kangxi-radical reorder)', () => {
    // gt sample: a 4-char CJK phrase whose 1st/3rd chars are Kangxi radicals (U+2F00/U+2F2F) drawn
    // out of content order; PDFium inserts generated spaces at the regressions
    const seq = [
      mkChar('⼀', 72, { fontSize: 15 }),
      mkChar(' ', 87, { fontSize: 15, width: 0, isGenerated: true }),
      mkChar('次', 87, { fontSize: 15 }),
      mkChar(' ', 102, { fontSize: 15, width: 0, isGenerated: true }),
      mkChar('⼯', 102, { fontSize: 15 }),
      mkChar('程', 117, { fontSize: 15 }),
    ]
    const words = groupIntoWords(seq)
    expect(texts(words)).toEqual(['⼀', '次', '⼯', '程'])
    expect(words.every((w) => !w.spaceBefore)).toBe(true)
  })

  it('drops a generated space at a CJK↔Latin boundary but keeps Latin↔Latin ones', () => {
    const zh = mkChar('中', 72, { fontSize: 10 })
    const genAfterZh = mkChar(' ', 82, { fontSize: 10, width: 0, isGenerated: true })
    const latin = mkText('ab', 84, { fontSize: 10 })
    const genLatin = mkChar(' ', latin.endX, { fontSize: 10, width: 0, isGenerated: true })
    const tail = mkText('cd', latin.endX + 2, { fontSize: 10 })
    const words = groupIntoWords([zh, genAfterZh, ...latin.chars, genLatin, ...tail.chars])
    expect(texts(words)).toEqual(['中', 'ab', 'cd'])
    expect(words[1]!.spaceBefore).toBe(false)
    expect(words[2]!.spaceBefore).toBe(true)
  })

  it('keeps a generated space at a Latin↔CJK boundary spanning a wide gap (P20: tab witness)', () => {
    // "5.3<tab>CJK title" heading numbering: the tab exports no glyph, PDFium
    // generates a zero-width space; the 18pt gap is real layout whitespace
    const num = mkText('5.3', 72, { fontSize: 12 })
    const gen = mkChar(' ', num.endX + 4, { fontSize: 1, width: 0, isGenerated: true })
    const zh = [
      mkChar('高', num.endX + 18, { fontSize: 12 }),
      mkChar('线', num.endX + 30, { fontSize: 12 }),
    ]
    const words = groupIntoWords([...num.chars, gen, ...zh])
    expect(texts(words)).toEqual(['5.3', '高', '线'])
    expect(words[1]!.spaceBefore).toBe(true)
    expect(words[2]!.spaceBefore).toBe(false)
  })

  it('a real space glyph mixed with generated ones still counts (author intent wins)', () => {
    const a = mkChar('中', 72)
    const gen = mkChar(' ', 82, { width: 0, isGenerated: true })
    const real = mkChar(' ', 82, { width: 2.5 })
    const b = mkChar('文', 84.5)
    const words = groupIntoWords([a, gen, real, b])
    expect(words[1]!.spaceBefore).toBe(true)
  })
})

describe('letter-spaced display text (P10 C)', () => {
  it('detects uniform tracking gaps and suppresses inferred spaces', () => {
    // "TOAST" with a 6pt gap between every 5pt-wide glyph at 10.5pt
    const t = mkText('TOAST', 72, { fontSize: 10.5, width: 5, tracking: 6 })
    const words = groupIntoWords(t.chars)
    expect(texts(words)).toEqual(['TOAST'])
    expect(words[0]!.spaceBefore).toBe(false)
  })

  it('keeps real space glyphs as word gaps inside a letter-spaced line', () => {
    const a = mkText('A', 72, { fontSize: 10.5, width: 5, tracking: 6 })
    const sp = mkChar(' ', a.endX + 6, { fontSize: 10.5, width: 2.5 })
    const b = mkText('TOAST', a.endX + 6 + 2.5 + 6, { fontSize: 10.5, width: 5, tracking: 6 })
    const words = groupIntoWords([...a.chars, sp, ...b.chars])
    expect(texts(words)).toEqual(['A', 'TOAST'])
    expect(words[1]!.spaceBefore).toBe(true)
  })

  it('normal text with one wide word gap still infers spaces', () => {
    const a = mkText('ab', 72, { fontSize: 10 })
    const b = mkText('cd', a.endX + 3, { fontSize: 10 })
    expect(texts(groupIntoWords([...a.chars, ...b.chars]))).toEqual(['ab', 'cd'])
  })

  it('rejects non-uniform column-scale gaps (table cells)', () => {
    const a = mkText('one', 72, { fontSize: 10 })
    const b = mkText('two', 150, { fontSize: 10 })
    const c = mkText('three', 300, { fontSize: 10 })
    // gaps of ~63pt and ~135pt are columns, not tracking
    expect(texts(groupIntoWords([...a.chars, ...b.chars, ...c.chars])).length).toBe(3)
  })
})
