/** P32: cross-page paragraph stitching (mid-paragraph source page boundaries). */
import { describe, expect, it } from 'vitest'
import { stitchCrossPageParagraphs } from '../src/pipeline'
import type { IrPage, Line, Span, TextBlock } from '../src/ir'
import type { Rect } from '../src/geometry'

const H = 800
const mkLine = (x0: number, x1: number, topY: number, text = 'x', fontSize = 12): Line => {
  const box: Rect = { x0, x1, y0: topY - fontSize, y1: topY }
  const span: Span = {
    text,
    box,
    fontSize,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
  }
  return { spans: [span], box, baseline: topY - fontSize * 0.8, endsWithHyphen: false }
}
const mkBlock = (lines: Line[], over: Partial<TextBlock> = {}): TextBlock => ({
  kind: 'text',
  box: lines
    .map((l) => l.box)
    .reduce((a, b) => ({
      x0: Math.min(a.x0, b.x0),
      y0: Math.min(a.y0, b.y0),
      x1: Math.max(a.x1, b.x1),
      y1: Math.max(a.y1, b.y1),
    })),
  align: 'left',
  firstLineIndentPt: 0,
  dir: 'ltr',
  lines,
  ...over,
})
const mkPage = (index: number, blocks: TextBlock[]): IrPage =>
  ({
    index,
    widthPt: 600,
    heightPt: H,
    rotation: 0,
    blocks,
    degraded: false,
    scanned: false,
  }) as IrPage

/** near-full page: one paragraph spanning most of the height, full-width tail */
const fullTail = () => mkBlock([mkLine(50, 550, 760), mkLine(50, 550, 100), mkLine(50, 548, 80)])

describe('stitchCrossPageParagraphs (P32)', () => {
  it('stitches a mid-paragraph boundary and marks the page', () => {
    const prev = mkPage(0, [fullTail()])
    const cur = mkPage(1, [
      mkBlock([mkLine(50, 400, 760, 'cont')]),
      mkBlock([mkLine(50, 300, 700)]),
    ])
    stitchCrossPageParagraphs([prev, cur])
    expect(cur.flowsFromPrev).toBe(true)
    expect(prev.blocks[0]!.kind === 'text' && (prev.blocks[0] as TextBlock).lines).toHaveLength(4)
    expect(cur.blocks).toHaveLength(1)
  })

  it('keeps the break when the tail line is short (finished paragraph)', () => {
    const prev = mkPage(0, [
      mkBlock([mkLine(50, 550, 760), mkLine(50, 550, 100), mkLine(50, 200, 80)]),
    ])
    const cur = mkPage(1, [mkBlock([mkLine(50, 400, 760)])])
    stitchCrossPageParagraphs([prev, cur])
    expect(cur.flowsFromPrev).toBeUndefined()
  })

  it('keeps the break on an underfull earlier page', () => {
    const prev = mkPage(0, [mkBlock([mkLine(50, 550, 760), mkLine(50, 548, 740)])])
    const cur = mkPage(1, [mkBlock([mkLine(50, 400, 760)])])
    stitchCrossPageParagraphs([prev, cur])
    expect(cur.flowsFromPrev).toBeUndefined()
  })

  it('keeps the break when the continuation is indented (new paragraph)', () => {
    const prev = mkPage(0, [fullTail()])
    const cur = mkPage(1, [mkBlock([mkLine(50, 400, 760)], { firstLineIndentPt: 20 })])
    stitchCrossPageParagraphs([prev, cur])
    expect(cur.flowsFromPrev).toBeUndefined()
  })
})
