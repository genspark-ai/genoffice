/** Decorative-rule → paragraph-border mapping (P7, pure geometry). */
import { describe, expect, it } from 'vitest'
import { applyDecorBorders } from '../src/analyze'
import type { Rect } from '../src/geometry'
import type { PageSection, Stroke, TextBlock } from '../src/ir'

function block(box: Rect): TextBlock {
  return {
    kind: 'text',
    lines: [
      {
        spans: [
          {
            text: 'x',
            box,
            fontSize: 10,
            fontFamily: 'Helvetica',
            bold: false,
            italic: false,
            color: '000000',
            dir: 'ltr',
            script: 'latin',
          },
        ],
        box,
        baseline: box.y0 + 2,
        endsWithHyphen: false,
      },
    ],
    box,
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
  }
}

function section(blocks: TextBlock[], box: Rect): PageSection {
  return { box, columns: [{ box, blocks }], gutterWidthsPt: [], dir: 'ltr' }
}

function hStroke(x0: number, x1: number, y: number, widthPt = 1.5, color = '0A3C61'): Stroke {
  return {
    box: { x0, x1, y0: y - widthPt / 2, y1: y + widthPt / 2 },
    orientation: 'h',
    widthPt,
    color,
  }
}

const COL: Rect = { x0: 54, y0: 100, x1: 530, y1: 800 }

describe('applyDecorBorders', () => {
  it('attaches a rule above a title as its w:top border (color, thickness, gap, right inset)', () => {
    const title = block({ x0: 54, y0: 763, x1: 138, y1: 782 })
    const sections = [section([title], COL)]
    const result = applyDecorBorders(sections, [hStroke(54, 189, 793.65)], [], new Set())
    expect(result.ignoredVertical).toBe(0)
    expect(title.border).toBeDefined()
    expect(title.border!.side).toBe('top')
    expect(title.border!.color).toBe('0A3C61')
    expect(title.border!.widthPt).toBeCloseTo(1.5)
    expect(title.border!.spacePt).toBeCloseTo(11.65, 1)
    // the rule reaches past the text → right inset pins the border at x=189
    expect(title.border!.indentRightPt).toBeCloseTo(530 - 189)
  })

  it('attaches a rule below a paragraph as its w:bottom border', () => {
    const para = block({ x0: 54, y0: 700, x1: 400, y1: 712 })
    const sections = [section([para], COL)]
    applyDecorBorders(sections, [hStroke(54, 400, 692)], [], new Set())
    expect(para.border?.side).toBe('bottom')
    expect(para.border?.spacePt).toBeCloseTo(8)
  })

  it('does not pin the right inset when the rule is narrower than its text', () => {
    const para = block({ x0: 54, y0: 700, x1: 400, y1: 712 })
    applyDecorBorders([section([para], COL)], [hStroke(54, 200, 692)], [], new Set())
    expect(para.border?.side).toBe('bottom')
    expect(para.border?.indentRightPt).toBeUndefined()
  })

  it('leaves table-region strokes and consumed underlines alone', () => {
    const para = block({ x0: 54, y0: 700, x1: 400, y1: 712 })
    const inTable = hStroke(60, 200, 695)
    const underline = hStroke(54, 400, 698)
    const sections = [section([para], COL)]
    applyDecorBorders(
      sections,
      [inTable, underline],
      [{ x0: 50, y0: 650, x1: 250, y1: 720 }],
      new Set([underline]),
    )
    expect(para.border).toBeUndefined()
    expect(sections[0]!.columns[0]!.blocks).toHaveLength(1)
  })

  it('with no host in range the rule becomes a standalone bordered bar paragraph', () => {
    const far = block({ x0: 54, y0: 100, x1: 400, y1: 112 }) // 300pt away
    const sections = [section([far], COL)]
    applyDecorBorders(sections, [hStroke(154, 430, 500)], [], new Set())
    const blocks = sections[0]!.columns[0]!.blocks
    expect(blocks).toHaveLength(2)
    const bar = blocks[0] as TextBlock // above `far` in flow (higher y first)
    expect(bar.kind).toBe('text')
    expect(bar.lines).toHaveLength(0)
    expect(bar.border?.side).toBe('top')
    expect(bar.border?.indentLeftPt).toBeCloseTo(100)
    expect(bar.border?.indentRightPt).toBeCloseTo(100)
  })

  it('counts vertical decorative strays and skips short/thick lines', () => {
    const para = block({ x0: 54, y0: 700, x1: 400, y1: 712 })
    const vertical: Stroke = {
      box: { x0: 300, x1: 301, y0: 400, y1: 600 },
      orientation: 'v',
      widthPt: 1,
      color: '000000',
    }
    const short = hStroke(54, 74, 692) // 20pt < 30pt minimum
    const thick = hStroke(54, 400, 690, 5) // 5pt > 3pt maximum
    const sections = [section([para], COL)]
    const result = applyDecorBorders(sections, [vertical, short, thick], [], new Set())
    expect(result.ignoredVertical).toBe(1)
    expect(para.border).toBeUndefined()
    expect(sections[0]!.columns[0]!.blocks).toHaveLength(1)
  })
})
