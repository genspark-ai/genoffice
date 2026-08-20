/** Double-drawn char dedup unit tests (P11 B): no wasm. */
import { describe, expect, it } from 'vitest'
import { dedupeDoubleDrawnChars, normalizeRegionalFontArtifacts } from '../src/analyze/chars'
import type { PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

describe('normalizeRegionalFontArtifacts (P16 D)', () => {
  const scText =
    '为加强全区各级行政机关政策解读工作规范政策执行强化政策落地保障公众权益办法说明这读们还进关问间'
  const mk = (text: string, fontFamily: string): PdfChar[] =>
    [...text].map((ch, i) => mkChar(ch, 100 + i * 12, { fontFamily, fontSize: 12 }))

  it('rewrites a TC family on strictly simplified text to FangSong', () => {
    const chars = mk(scText, 'Yuanti TC')
    normalizeRegionalFontArtifacts(chars)
    expect(chars.every((c) => c.fontFamily === 'FangSong')).toBe(true)
  })

  it('maps shape classes: Songti TC → SimSun, Heiti TC → SimHei', () => {
    const chars = [...mk(scText, 'Songti TC'), ...mk(scText, 'Heiti TC')]
    normalizeRegionalFontArtifacts(chars)
    expect(chars.slice(0, 5).every((c) => c.fontFamily === 'SimSun')).toBe(true)
    expect(chars.slice(-5).every((c) => c.fontFamily === 'SimHei')).toBe(true)
  })

  it('leaves TC families on traditional text alone', () => {
    const chars = mk(
      '為加強全區各級行政機關政策解讀工作規範這讀們還進關問間門業務經條現發見車書長',
      'Yuanti TC',
    )
    normalizeRegionalFontArtifacts(chars)
    expect(chars.every((c) => c.fontFamily === 'Yuanti TC')).toBe(true)
  })

  it('needs enough simplified evidence before rewriting', () => {
    const chars = mk('为了', 'Yuanti TC') // only one indicator char
    normalizeRegionalFontArtifacts(chars)
    expect(chars.every((c) => c.fontFamily === 'Yuanti TC')).toBe(true)
  })

  it('never touches non-TC families', () => {
    const chars = mk(scText, 'SimSun')
    normalizeRegionalFontArtifacts(chars)
    expect(chars.every((c) => c.fontFamily === 'SimSun')).toBe(true)
  })
})

describe('dedupeDoubleDrawnChars', () => {
  it('collapses a re-stroked twin onto the later-drawn glyph', () => {
    const first = mkChar('赴', 100, { y: 400, fontSize: 36 })
    const ghost = mkChar('赴', 101.5, { y: 401, fontSize: 36 }) // ~1.8pt off, same size
    const chars = [first, ghost]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(1)
    expect(chars[0]).toBe(ghost)
  })

  it('keeps the bigger glyph when the earlier twin is larger', () => {
    const big = mkChar('日', 100, { y: 400, fontSize: 36 })
    const small = mkChar('日', 101, { y: 400.5, fontSize: 30 }) // ratio 0.83, dist < 0.3 em
    const chars = [big, small]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(1)
    expect(chars[0]).toBe(big)
  })

  it('leaves normal tight-set repeated letters alone', () => {
    const chars = mkText('ll', 100, { fontSize: 10 }).chars // advance 0.5 em apart
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('leaves narrow letters under the center gate but with disjoint ink alone', () => {
    // real-font 'll': ~0.22 em advance (centers pass the 0.3 em gate) but the
    // thin ink boxes never overlap
    const chars = [
      mkChar('l', 100, { fontSize: 10, width: 1.8 }),
      mkChar('l', 102.2, { fontSize: 10, width: 1.8 }),
    ]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('leaves different characters overlapping (marks, ligature debris) alone', () => {
    const chars = [mkChar('e', 100, { fontSize: 10 }), mkChar('é', 100.2, { fontSize: 10 })]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('twins beyond the size band are not merged (heading vs body)', () => {
    const chars = [
      mkChar('A', 100, { y: 400, fontSize: 24 }),
      mkChar('A', 101, { y: 401, fontSize: 12 }),
    ]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('keeps both chars of a ligature glyph expanded through ToUnicode', () => {
    // a 'tt' ligature: PDFium copies the single glyph's box verbatim onto
    // both expanded chars ("Bottom" must not become "Botom")
    const first = mkChar('t', 84.035, { y: 724, fontSize: 11, width: 6.6 })
    const twin = mkChar('t', 84.035, { y: 724, fontSize: 11, width: 6.6 })
    const chars = [first, twin]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('still collapses identical-box twins that are not stream-adjacent', () => {
    const first = mkChar('赴', 100, { y: 400, fontSize: 36 })
    const other = mkChar('日', 200, { y: 400, fontSize: 36 })
    const ghost = mkChar('赴', 100, { y: 400, fontSize: 36 })
    const chars = [first, other, ghost]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(2)
  })

  it('collapses a triple strike to a single survivor', () => {
    const chars: PdfChar[] = [
      mkChar('業', 200, { y: 300, fontSize: 28 }),
      mkChar('業', 200.8, { y: 300.4, fontSize: 28 }),
      mkChar('業', 201.4, { y: 300.9, fontSize: 28 }),
    ]
    dedupeDoubleDrawnChars(chars)
    expect(chars).toHaveLength(1)
  })
})
