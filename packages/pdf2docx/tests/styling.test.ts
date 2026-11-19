import { describe, expect, it } from 'vitest'
import { analyzeChars, applyTextShapeStyles } from '../src/analyze'
import type { Fill, PageShapes, Stroke } from '../src/ir'
import { nearestHighlight } from '../src/rebuild'
import { mkText } from './helpers/chars'

const PAGE_AREA = 612 * 792

const hStroke = (x0: number, x1: number, y: number, widthPt = 1, color = '000000'): Stroke => ({
  box: { x0, x1, y0: y - widthPt / 2, y1: y + widthPt / 2 },
  orientation: 'h',
  widthPt,
  color,
})

const fill = (x0: number, y0: number, x1: number, y1: number, color: string): Fill => ({
  box: { x0, y0, x1, y1 },
  color,
})

const shapesOf = (strokes: Stroke[] = [], fills: Fill[] = []): PageShapes => ({
  strokes,
  fills,
  ignoredPaths: 0,
})

describe('applyTextShapeStyles: highlights', () => {
  it('marks chars covered by a non-white fill with the fill color', () => {
    const { chars } = mkText('mark', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([], [fill(98, 697, 122, 709, 'FFFF00')])
    const styled = applyTextShapeStyles(chars, shapes, [], PAGE_AREA)
    expect(styled.chars.every((c) => c.highlight === 'FFFF00')).toBe(true)
  })

  it('ignores white fills, page-sized fills and fills inside table regions', () => {
    const { chars } = mkText('mark', 100, { y: 700, fontSize: 10 })
    const white = shapesOf([], [fill(98, 697, 122, 709, 'FFFFFF')])
    expect(applyTextShapeStyles(chars, white, [], PAGE_AREA).chars.some((c) => c.highlight)).toBe(
      false,
    )

    const pageBg = shapesOf([], [fill(0, 0, 612, 792, 'FFFF00')])
    expect(applyTextShapeStyles(chars, pageBg, [], PAGE_AREA).chars.some((c) => c.highlight)).toBe(
      false,
    )

    const shaded = shapesOf([], [fill(98, 697, 122, 709, 'FFFF00')])
    const tableBox = { x0: 90, y0: 690, x1: 200, y1: 720 }
    expect(
      applyTextShapeStyles(chars, shaded, [tableBox], PAGE_AREA).chars.some((c) => c.highlight),
    ).toBe(false)
  })

  it('ignores tall text-box backgrounds', () => {
    const { chars } = mkText('boxed', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([], [fill(90, 640, 200, 760, '00FF00')])
    expect(applyTextShapeStyles(chars, shapes, [], PAGE_AREA).chars.some((c) => c.highlight)).toBe(
      false,
    )
  })
})

describe('applyTextShapeStyles: underline and strikethrough', () => {
  it('maps a baseline-hugging stroke to underline', () => {
    const { chars, endX } = mkText('under', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([hStroke(99, endX + 1, 699)])
    const styled = applyTextShapeStyles(chars, shapes, [], PAGE_AREA)
    expect(styled.chars.every((c) => c.underline === true)).toBe(true)
    expect(styled.chars.some((c) => c.strike)).toBe(false)
  })

  it('maps a mid-band stroke to strikethrough', () => {
    const { chars, endX } = mkText('gone', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([hStroke(99, endX + 1, 703.5)])
    const styled = applyTextShapeStyles(chars, shapes, [], PAGE_AREA)
    expect(styled.chars.every((c) => c.strike === true)).toBe(true)
    expect(styled.chars.some((c) => c.underline)).toBe(false)
  })

  it('warns when the line color differs from the text color', () => {
    const { chars, endX } = mkText('red', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([hStroke(99, endX + 1, 699, 1, 'FF0000')])
    const styled = applyTextShapeStyles(chars, shapes, [], PAGE_AREA)
    expect(styled.chars.every((c) => c.underline === true)).toBe(true)
    expect(styled.warnings.length).toBe(1)
  })

  it('does not style text from table borders or long separator rules', () => {
    const { chars, endX } = mkText('cell', 100, { y: 700, fontSize: 10 })
    // stroke inside a lattice region → excluded outright
    const border = shapesOf([hStroke(99, endX + 1, 699)])
    const tableBox = { x0: 90, y0: 690, x1: 200, y1: 720 }
    expect(
      applyTextShapeStyles(chars, border, [tableBox], PAGE_AREA).chars.some((c) => c.underline),
    ).toBe(false)

    // a full-width rule under a short word is a separator, not an underline
    const separator = shapesOf([hStroke(72, 540, 699)])
    expect(
      applyTextShapeStyles(chars, separator, [], PAGE_AREA).chars.some((c) => c.underline),
    ).toBe(false)
  })

  it('ignores thick decorative bars', () => {
    const { chars, endX } = mkText('bar', 100, { y: 700, fontSize: 10 })
    const shapes = shapesOf([hStroke(99, endX + 1, 699, 6)])
    expect(applyTextShapeStyles(chars, shapes, [], PAGE_AREA).chars.some((c) => c.underline)).toBe(
      false,
    )
  })
})

describe('styled chars flow into spans', () => {
  it('splits spans at style-flag boundaries and carries the flags', () => {
    const plain = mkText('plain ', 100, { y: 700, fontSize: 10 })
    const marked = mkText('marked', plain.endX + 2.5, { y: 700, fontSize: 10 })
    const shapes = shapesOf(
      [hStroke(plain.endX + 1.5, plain.endX + 35, 699)],
      [fill(plain.endX + 1.5, 697, plain.endX + 35, 709, 'FFFF00')],
    )
    const styled = applyTextShapeStyles([...plain.chars, ...marked.chars], shapes, [], PAGE_AREA)
    const lines = analyzeChars(styled.chars)
    expect(lines).toHaveLength(1)
    const spans = lines[0]!.spans
    expect(spans.length).toBe(2)
    expect(spans[0]!.underline).toBeUndefined()
    expect(spans[1]!.underline).toBe(true)
    expect(spans[1]!.highlight).toBe('FFFF00')
  })
})

describe('nearestHighlight', () => {
  it('maps hex colors to the closest named w:highlight', () => {
    expect(nearestHighlight('FFFF00')).toBe('yellow')
    expect(nearestHighlight('FFD100')).toBe('yellow')
    expect(nearestHighlight('00E5EE')).toBe('cyan')
    expect(nearestHighlight('CC0011')).toBe('red')
  })

  it('treats near-white as "no highlight"', () => {
    expect(nearestHighlight('FFFFFF')).toBeUndefined()
    expect(nearestHighlight('FAFAF7')).toBeUndefined()
  })
})

describe('display-size glyphs never take highlights (P10 C)', () => {
  it('skips w:highlight for a stamp-sized char over a dark fill', () => {
    // a 135pt monogram over a near-black glow slab must not become a black slab
    const { chars } = mkText('敬', 400, { y: 300, fontSize: 135 })
    const shapes = shapesOf([], [fill(369, 189, 591, 411, '000000')])
    expect(applyTextShapeStyles(chars, shapes, [], 960 * 540).chars.some((c) => c.highlight)).toBe(
      false,
    )
  })
})
