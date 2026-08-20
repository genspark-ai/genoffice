/** before_space chain unit tests: hand-built sections, no wasm. */
import { describe, expect, it } from 'vitest'
import { applySpacingChain } from '../src/analyze/spacing'
import type { Line, PageColumn, PageSection, Span, TextBlock } from '../src/ir'

function span(text: string, x0 = 72, x1 = 300): Span {
  return {
    text,
    box: { x0, y0: 0, x1, y1: 0 },
    fontSize: 12,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
  }
}

/** a 12pt-tall line whose ink box top sits at `top` */
function line(top: number, text = 'line'): Line {
  return {
    spans: [span(text)],
    box: { x0: 72, x1: 300, y0: top - 12, y1: top },
    baseline: top - 10,
    endsWithHyphen: false,
  }
}

/** block from line tops (each line 12pt tall) */
function block(...tops: number[]): TextBlock {
  const lines = tops.map((t) => line(t))
  return {
    kind: 'text',
    lines,
    box: {
      x0: 72,
      x1: 300,
      y0: Math.min(...lines.map((l) => l.box.y0)),
      y1: Math.max(...lines.map((l) => l.box.y1)),
    },
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
  }
}

function sectionOf(columns: PageColumn[]): PageSection {
  const boxes = columns.flatMap((c) => c.blocks.map((b) => b.box))
  return {
    box: {
      x0: Math.min(...boxes.map((b) => b.x0)),
      x1: Math.max(...boxes.map((b) => b.x1)),
      y0: Math.min(...boxes.map((b) => b.y0)),
      y1: Math.max(...boxes.map((b) => b.y1)),
    },
    columns,
    gutterWidthsPt: [],
    dir: 'ltr',
  }
}

const column = (...blocks: TextBlock[]): PageColumn => ({
  box: { x0: 72, x1: 300, y0: 0, y1: 792 },
  blocks,
})

describe('applySpacingChain', () => {
  it('turns the full inter-block ink gap into spacingBefore (paragraphs occupy exact ink extents)', () => {
    // a: lines 700/686 (bottom 674); b: top 640 → gap 34
    const a = block(700, 686)
    const b = block(640)
    applySpacingChain([sectionOf([column(a, b)])])
    expect(a.spacingBeforePt).toBeUndefined() // page-top leading is rebuild's job
    expect(b.spacingBeforePt).toBeCloseTo(34, 5)
  })

  it('leaves sub-threshold spacing unset', () => {
    // gap 688−686.5 = 1.5pt < 2pt noise floor
    const a = block(700)
    const b = block(686.5)
    applySpacingChain([sectionOf([column(a, b)])])
    expect(b.spacingBeforePt).toBeUndefined()
  })

  it('clamps negative gaps to 0 and warns', () => {
    const a = block(700)
    const b = block(695) // overlaps a (a bottom = 688 > b top − …)
    const warnings = applySpacingChain([sectionOf([column(a, b)])])
    expect(b.spacingBeforePt).toBeUndefined()
    expect(warnings).toHaveLength(1)
    expect(warnings[0]).toMatch(/negative gap/)
  })

  it('chains the first block of a later section from the previous section bottom', () => {
    const a = block(700, 686) // section 1, bottom 674
    const b = block(600) // section 2 first block, gap 74
    applySpacingChain([sectionOf([column(a)]), sectionOf([column(b)])])
    expect(b.spacingBeforePt).toBeCloseTo(74, 5)
  })

  it('gives columns after the first no section-top spacing', () => {
    const head = block(750, 736) // leading sample
    const left = block(700)
    const right = block(700)
    const twoCol = sectionOf([column(left), column(right)])
    applySpacingChain([sectionOf([column(head)]), twoCol])
    expect(left.spacingBeforePt).toBeDefined() // reading-order first block chains
    expect(right.spacingBeforePt).toBeUndefined() // column 2 starts mid-flow
  })

  it('skips floats entirely', () => {
    const a = block(700)
    const float: import('../src/ir').ImageBlock = {
      kind: 'image' as const,
      box: { x0: 100, y0: 660, x1: 200, y1: 680 },
      data: new Uint8Array(0),
      mime: 'image/png' as const,
      pixelWidth: 1,
      pixelHeight: 1,
      float: { wrap: 'square-left' as const, xOffsetPt: 28 },
    }
    const b = block(640)
    const col = column(a, b)
    col.blocks.splice(1, 0, float)
    applySpacingChain([sectionOf([col])])
    // chain from a (bottom 688) to b (top 640): full 48pt gap
    expect(b.spacingBeforePt).toBeCloseTo(48, 5)
    expect(float.spacingBeforePt).toBeUndefined()
  })
})
