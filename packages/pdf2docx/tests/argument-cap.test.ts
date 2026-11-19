/**
 * V8's spread-argument limit regression: `Math.min(...pageSizedArray)` passes
 * every element as a CALL ARGUMENT, so a page past ~124,600 units threw
 * RangeError inside analyze/rebuild and aborted the whole PDF->DOCX run
 * (analyzePage and the extractIrDocument page loop have no try/catch).
 * Every page-sized reduction must be a loop. Measured on V8: 100k args OK,
 * 125k args throws.
 *
 * These drive the pure-logic entry points directly, so they need no wasm.
 */
import { describe, expect, it } from 'vitest'
import { detectSections } from '../src/analyze/columns'
import { detectFootnotes } from '../src/analyze/footnotes'
import { classifyFloatImages } from '../src/analyze/floats'
import type { ImageBlock, PageShapes, PdfChar } from '../src/ir'
import type { LineUnit } from '../src/analyze/units'
import type { Rect } from '../src/geometry'

/** well past the ~124,600-argument cap */
const OVER_CAP = 130_000

const mkChar = (i: number): PdfChar => {
  const x0 = 72 + (i % 50)
  return {
    code: 65,
    text: 'A',
    box: { x0, x1: x0 + 20, y0: 698, y1: 710 },
    looseBox: { x0, x1: x0 + 20, y0: 696, y1: 712 },
    originX: x0,
    originY: 700,
    angle: 0,
    fontSize: 10,
    fontWeight: 400,
    fontFamily: 'Helvetica',
    italic: false,
    color: '000000',
    isGenerated: false,
    isHyphen: false,
    script: 'latin',
  }
}

/**
 * A line unit per glyph. The last one sits in a far band so the page's y span
 * (min y0 .. max y1) is much taller than the text band: that is what lets a
 * tiny icon land INSIDE the page's text span without any unit overlapping it,
 * so classifyFloatImages reaches its TINY_ICON branch.
 */
const mkUnit = (i: number): LineUnit => {
  const band = i === OVER_CAP - 1 ? { y0: 90, y1: 95 } : { y0: 698, y1: 710 }
  const x0 = 72 + (i % 50)
  return {
    chars: [mkChar(i)],
    box: { x0, x1: x0 + 20, ...band },
    baseline: band.y1,
    lineIndex: i,
    wordCount: 1,
    fontSize: 10,
  }
}

describe('page-sized reductions past the V8 argument cap', () => {
  it('classifies a tiny icon on a 130k-unit page without throwing', () => {
    const units = Array.from({ length: OVER_CAP }, (_, i) => mkUnit(i))
    // 10x10pt, inside the page's y span, clear of every unit in x and y
    const icon: ImageBlock = {
      kind: 'image',
      box: { x0: 300, x1: 310, y0: 200, y1: 210 },
      data: new Uint8Array([0]),
      mime: 'image/png',
      pixelWidth: 1,
      pixelHeight: 1,
    }

    // reached the TINY_ICON branch (behind-anchored) AND completed: the y
    // extents used to be `Math.max(...units.map(...))` inside this branch
    const { floats, inline } = classifyFloatImages([icon], units)
    expect(floats).toHaveLength(1)
    expect(floats[0]!.float).toEqual({ wrap: 'behind', xOffsetPt: 300 - 72 })
    expect(inline).toHaveLength(0)
  })

  it('detects footnotes on a 130k-glyph page that has a stroke', () => {
    const chars = Array.from({ length: OVER_CAP }, (_, i) => mkChar(i))
    const shapes: PageShapes = {
      strokes: [
        {
          orientation: 'h',
          widthPt: 200,
          box: { x0: 72, x1: 272, y0: 100, y1: 101 },
          color: '000000',
        },
      ],
      fills: [],
      ignoredPaths: 0,
    }

    // contentLeft used to be `Math.min(...visible.map(...))` over every visible
    // glyph, guarded only by a non-empty check and one stroke
    const found = detectFootnotes(chars, shapes, 0, 612, 792)
    expect(found.footnotes).toEqual([])
    expect(found.bodyChars).toHaveLength(OVER_CAP)
  })

  it('detects sections over a page-sized element array', () => {
    const elements = Array.from({ length: OVER_CAP }, (_, i) => ({
      box: { x0: 72, x1: 92, y0: 698, y1: 710 } as Rect,
      unit: mkUnit(i),
    }))

    // contentWidthOf/heightOf used to spread col.elements; detectSections is
    // the only way in, and isStrong re-measures every column per gutter
    const sections = detectSections(elements, 792, 612)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(1)
  })
})
