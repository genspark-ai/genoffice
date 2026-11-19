import { describe, expect, it } from 'vitest'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze'
import { mkChar, mkText } from './helpers/chars'

const lineTexts = (lines: ReturnType<typeof groupIntoLines>) =>
  lines.map((l) => l.chars.map((c) => c.text).join(''))

describe('baseline line clustering', () => {
  it('keeps same-baseline chars on one line', () => {
    const { chars } = mkText('one line', 72, { y: 700 })
    expect(groupIntoLines(chars)).toHaveLength(1)
  })

  it('splits on a baseline change beyond the fontSize-relative tolerance', () => {
    const l1 = mkText('first', 72, { y: 700, fontSize: 10 })
    const l2 = mkText('second', 72, { y: 688, fontSize: 10 }) // 12pt drop >> 4.5pt tol
    const lines = groupIntoLines([...l1.chars, ...l2.chars])
    expect(lineTexts(lines)).toEqual(['first', 'second'])
  })

  it('tolerates small baseline jitter relative to the font size', () => {
    const a = mkText('wob', 72, { y: 700, fontSize: 10 })
    const b = mkText('ble', a.endX, { y: 701.5, fontSize: 10 }) // 1.5pt < 4.5pt tol
    expect(groupIntoLines([...a.chars, ...b.chars])).toHaveLength(1)
  })

  it('scales the tolerance with the font size (same jitter, big type)', () => {
    const a = mkText('BIG', 72, { y: 700, fontSize: 36 })
    const b = mkText('GER', a.endX, { y: 708, fontSize: 36 }) // 8pt < 0.45×36
    expect(groupIntoLines([...a.chars, ...b.chars])).toHaveLength(1)
  })

  it('splits on generated newline characters when the baseline moves', () => {
    const l1 = mkText('para', 72, { y: 700 })
    const nl = mkChar('\n', l1.endX, { y: 700, isGenerated: true, width: 0 })
    const l2 = mkText('graph', 72, { y: 688 })
    const lines = groupIntoLines([...l1.chars, nl, ...l2.chars])
    expect(lineTexts(lines)).toEqual(['para', 'graph'])
  })

  it('ignores generated newlines that stay on the baseline (letter-spaced titles, P10 C)', () => {
    // PDFium fabricates a newline between separately-positioned glyphs of a
    // spaced-out title; the run continues forward on the same baseline
    const a = mkText('TO', 72, { y: 700 })
    const nl = mkChar('\n', a.endX, { y: 700, isGenerated: true, width: 0 })
    const b = mkText('AST', a.endX + 8, { y: 700 })
    expect(lineTexts(groupIntoLines([...a.chars, nl, ...b.chars]))).toEqual(['TOAST'])
  })

  it('still splits on a real newline glyph on the same baseline', () => {
    const a = mkText('para', 72, { y: 700 })
    const nl = mkChar('\n', a.endX, { y: 700, isGenerated: false, width: 0 })
    const b = mkText('graph', a.endX + 8, { y: 700 })
    expect(lineTexts(groupIntoLines([...a.chars, nl, ...b.chars]))).toEqual(['para', 'graph'])
  })

  it('splits on a hard x regression (new visual line without a newline char)', () => {
    const l1 = mkText('right side text', 200, { y: 700 })
    const l2 = mkText('back left', 72, { y: 700 })
    const lines = groupIntoLines([...l1.chars, ...l2.chars])
    expect(lines).toHaveLength(2)
  })

  it('keeps superscripts on the line via vertical box overlap', () => {
    const base = mkText('E=mc', 72, { y: 700, fontSize: 10 })
    const sup = mkChar('2', base.endX, { y: 704, fontSize: 6 }) // raised baseline, box overlaps
    const lines = groupIntoLines([...base.chars, sup])
    expect(lines).toHaveLength(1)
    expect(lineTexts(lines)).toEqual(['E=mc2'])
  })

  it('flags end-of-line hyphenation', () => {
    const l1 = mkText('hyphen', 72, { y: 700 })
    const dash = mkChar('-', l1.endX, { y: 700, isHyphen: true, width: 3 })
    const l2 = mkText('ated', 72, { y: 688 })
    const lines = groupIntoLines([...l1.chars, dash, ...l2.chars])
    expect(lines[0]!.endsWithHyphen).toBe(true)
    expect(lines[1]!.endsWithHyphen).toBe(false)
  })

  it('computes the line baseline as the median char origin', () => {
    const a = mkText('abc', 72, { y: 700 })
    const lines = groupIntoLines(a.chars)
    expect(lines[0]!.baseline).toBeCloseTo(700)
  })
})

describe('combining mark clustering', () => {
  it('merges combining marks into the preceding base char', () => {
    const e = mkChar('e', 72)
    const acute = mkChar('́', 74, { width: 0 }) // zero-width overlay
    const x = mkChar('x', 77)
    const clustered = clusterCombiningMarks([e, acute, x])
    expect(clustered).toHaveLength(2)
    expect(clustered[0]!.text).toBe('é')
  })

  it('merges Thai vowel/tone marks into the base consonant cluster', () => {
    const base = mkChar('ก', 72, { width: 6 })
    const vowel = mkChar('ิ', 74, { width: 0 })
    const tone = mkChar('่', 74.5, { width: 0 })
    const clustered = clusterCombiningMarks([base, vowel, tone])
    expect(clustered).toHaveLength(1)
    expect(clustered[0]!.text).toBe('กิ่')
  })

  it('keeps a leading mark standalone (no base to attach to)', () => {
    const mark = mkChar('́', 72, { width: 0 })
    const a = mkChar('a', 74)
    expect(clusterCombiningMarks([mark, a])).toHaveLength(2)
  })
})

describe('hyphen-marker tail (P21 B)', () => {
  it('flags hyphenation when PDFium replaced the glyph with the 0x02 marker', () => {
    const l1 = mkText('ex', 72, { y: 700 })
    // PDFium's replacement marker: control code, filtered from visible chars
    const marker = mkChar('\u0002', l1.endX, { y: 700, isHyphen: true, width: 2.5 })
    const l2 = mkText('pected', 72, { y: 688 })
    const lines = groupIntoLines([...l1.chars, marker, ...l2.chars])
    expect(lines[0]!.endsWithHyphen).toBe(true)
    expect(lines[1]!.endsWithHyphen).toBe(false)
  })
})
